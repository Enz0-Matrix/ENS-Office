use chrono::Datelike;

use crate::{
    locale::Locale,
    number_format::{parse_finite_number, to_precision},
};

use super::{
    dates::{date_to_serial_number, from_excel_date},
    parser::{ParsePart, Parser, TextToken},
};

pub struct Formatted {
    pub color: Option<i32>,
    pub text: String,
    pub error: Option<String>,
}

/// Returns the vector of chars of the fractional part of a *positive* number:
/// 3.1415926 ==> ['1', '4', '1', '5', '9', '2', '6']
fn get_fract_part(value: f64, precision: i32, int_len: usize) -> Vec<char> {
    let b = format!("{:.1$}", value.fract(), precision as usize)
        .chars()
        .collect::<Vec<char>>();
    let l = b.len() - 1;
    let mut last_non_zero = b.len() - 1;
    for i in 0..l {
        if b[l - i] != '0' {
            last_non_zero = l - i + 1;
            break;
        }
    }
    if last_non_zero < 2 {
        return vec![];
    }
    let max_len = if int_len > 15 {
        2_usize
    } else {
        15_usize - int_len + 1
    };
    let last_non_zero = usize::min(last_non_zero, max_len + 1);
    b[2..last_non_zero].to_vec()
}

/// Return true if we need to add a separator in position digit_index
/// It normally happens if if digit_index -1 is 3, 6, 9,... digit_index ≡ 1 mod 3
fn use_group_separator(use_thousands: bool, digit_index: i32, group_sizes: &str) -> bool {
    if use_thousands {
        if group_sizes == "#,##0.###" {
            if digit_index > 1 && (digit_index - 1) % 3 == 0 {
                return true;
            }
        } else if group_sizes == "#,##,##0.###"
            && (digit_index == 3 || (digit_index > 3 && digit_index % 2 == 0))
        {
            return true;
        }
    }
    false
}

/// ENS-Office: the genitive month name ("марта") for locales that need it
/// after a day number; `None` for the others.
fn genitive_month(locale: &Locale, month: usize) -> Option<&'static str> {
    const RU: [&str; 12] = [
        "января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря",
    ];
    if locale.dates.months.first().map(|m| m.to_lowercase()) == Some("январь".to_string()) {
        RU.get(month.checked_sub(1)?).copied()
    } else {
        None
    }
}

/// ENS-Office: round to `digits` decimals, halves away from zero, working on
/// the shortest decimal representation (so 2.345 is "2.345", not 2.34499…).
pub(crate) fn round_half_away(value: f64, digits: i32) -> f64 {
    if !value.is_finite() || !(0..=15).contains(&digits) {
        return value;
    }
    let clean = to_precision(value, 15);
    let s = format!("{}", clean.abs());
    let (int_part, frac_part) = match s.split_once('.') {
        Some((i, f)) => (i, f),
        None => return clean,
    };
    let d = digits as usize;
    if frac_part.len() <= d {
        return clean;
    }
    let mut kept: Vec<u8> = int_part.bytes().chain(frac_part.bytes().take(d)).map(|b| b - b'0').collect();
    if frac_part.as_bytes()[d] >= b'5' {
        let mut i = kept.len();
        loop {
            if i == 0 {
                kept.insert(0, 1);
                break;
            }
            i -= 1;
            if kept[i] == 9 {
                kept[i] = 0;
            } else {
                kept[i] += 1;
                break;
            }
        }
    }
    let digits_str: String = kept.iter().map(|d| (b'0' + d) as char).collect();
    let split = digits_str.len() - d;
    let text = if d == 0 { digits_str } else { format!("{}.{}", &digits_str[..split], &digits_str[split..]) };
    let rounded = text.parse::<f64>().unwrap_or(clean);
    if clean < 0.0 { -rounded } else { rounded }
}

/// ENS-Office: a format code typed in a formula (TEXT) in the locale's notation —
/// "# ##0,00" in Russian Excel — to the stored notation "#,##0.00". Text in
/// quotes and escaped characters are left alone.
pub fn delocalize_format_code(code: &str, locale: &Locale) -> String {
    if locale.numbers.symbols.decimal != "," {
        return code.to_string();
    }
    let chars: Vec<char> = code.chars().collect();
    let mut out = String::with_capacity(code.len());
    let mut i = 0;
    let mut in_quotes = false;
    let is_digit = |c: Option<&char>| matches!(c, Some('0' | '#' | '?'));
    while i < chars.len() {
        let c = chars[i];
        if c == '"' {
            in_quotes = !in_quotes;
            out.push(c);
        } else if in_quotes {
            out.push(c);
        } else if c == '\\' && i + 1 < chars.len() {
            out.push(c);
            out.push(chars[i + 1]);
            i += 1;
        } else if c == ',' {
            out.push('.');
        } else if c == '.' {
            out.push(',');
        } else if (c == ' ' || c == '\u{a0}') && is_digit(chars.get(i.wrapping_sub(1)).filter(|_| i > 0)) && is_digit(chars.get(i + 1)) {
            // a space between digit placeholders is the thousands separator
            out.push(',');
        } else {
            out.push(c);
        }
        i += 1;
    }
    out
}

/// ENS-Office: the sections of a format code (split at `;` outside quotes).
fn format_sections(format: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let (mut start, mut quoted, mut escaped) = (0, false, false);
    for (i, c) in format.char_indices() {
        if escaped {
            escaped = false;
        } else if c == '\\' {
            escaped = true;
        } else if c == '"' {
            quoted = !quoted;
        } else if c == ';' && !quoted {
            out.push(&format[start..i]);
            start = i + 1;
        }
    }
    out.push(&format[start..]);
    out
}

/// ENS-Office: digits of a fraction part padded as Excel does: `?` → space, `0` → zero.
fn pad_fraction(digits: String, pattern: &str, left: bool) -> String {
    let n = pattern.chars().count();
    let len = digits.chars().count();
    if len >= n {
        return digits;
    }
    let mut pad = String::new();
    for c in pattern.chars().take(n - len) {
        match c {
            '?' => pad.push(' '),
            '0' if left => pad.push('0'),
            '0' => pad.push(' '),
            _ => {}
        }
    }
    if left {
        format!("{pad}{digits}")
    } else {
        format!("{digits}{pad}")
    }
}

/// ENS-Office: fraction formats as in Excel — `# ?/?`, `# ??/??`, `?/?`, `# ?/8`, `# ??/100`.
/// None when the code is not a fraction format (then the usual formatter runs).
fn format_fraction(value_original: f64, format: &str) -> Option<Formatted> {
    if !format.contains('/') {
        return None;
    }
    let sections = format_sections(format);
    let v = value_original;
    let (section, value, signed) = match sections.len() {
        1 => (sections[0], v, true),
        2 => if v >= 0.0 { (sections[0], v, false) } else { (sections[1], -v, false) },
        _ => {
            if v > 0.0 {
                (sections[0], v, false)
            } else if v < 0.0 {
                (sections[1], -v, false)
            } else {
                (sections[2], 0.0, false)
            }
        }
    };
    // [Цвет] / [условие] в начале раздела
    let mut body = section;
    let mut prefix = String::new();
    while body.starts_with('[') {
        let end = body.find(']')?;
        prefix.push_str(&body[..=end]);
        body = &body[end + 1..];
    }
    let (left, den) = body.split_once('/')?;
    let is_digit_pattern = |s: &str| !s.is_empty() && s.chars().all(|c| matches!(c, '?' | '#' | '0'));
    let fixed_den: Option<u64> = if den.chars().all(|c| c.is_ascii_digit()) && !den.starts_with('0') && !den.is_empty() {
        den.parse().ok()
    } else if is_digit_pattern(den) && den.len() <= 4 {
        None
    } else {
        return None;
    };
    let left_trim = left.trim_end_matches(['?', '#', '0']);
    let num_pattern = &left[left_trim.len()..];
    if num_pattern.is_empty() {
        return None;
    }
    let int_end = left_trim.trim_end();
    let int_pattern = &int_end[int_end.trim_end_matches(['?', '#', '0', ',']).len()..];
    let has_int = !int_pattern.is_empty();
    if has_int && int_end.len() == left_trim.len() {
        return None; // «#?/?» без пробела — не дробь
    }
    // текст перед числом («-», "ок. »)
    let lead_src = if has_int { &int_end[..int_end.len() - int_pattern.len()] } else { left_trim };
    let lead: String = lead_src.chars().filter(|c| !matches!(c, '"' | '\\')).collect();
    if lead.chars().any(|c| matches!(c, '?' | '#' | '0')) {
        return None;
    }
    let left_trim = &left_trim[lead_src.len()..];
    let int_pattern = if has_int { int_pattern } else { "" };

    let x = value.abs();
    let (mut int, frac) = if has_int { (x.floor(), x - x.floor()) } else { (0.0, x) };
    let (mut n, d) = match fixed_den {
        Some(d) => ((frac * d as f64).round() as u64, d),
        None => {
            let max_d = 10u64.pow(den.len() as u32) - 1;
            let (mut best_n, mut best_d, mut best_err) = (frac.round() as u64, 1u64, f64::MAX);
            for d in 1..=max_d {
                let n = (frac * d as f64).round();
                let err = (frac - n / d as f64).abs();
                if err < best_err - 1e-12 {
                    (best_n, best_d, best_err) = (n as u64, d, err);
                }
            }
            (best_n, best_d)
        }
    };
    if has_int && n >= d {
        int += (n / d) as f64;
        n %= d;
    }
    let minus = if signed && value < 0.0 && (int > 0.0 || n > 0) { "-" } else { "" };
    let gap = " ".repeat(left_trim.len() - int_pattern.len());
    let int_text = if int > 0.0 || int_pattern.contains('0') { format!("{int}") } else { String::new() };
    let text = if has_int && n == 0 {
        // целое: на месте дроби — пробелы (числа в столбце выравниваются)
        let width = num_pattern.chars().count() + 1 + den.chars().count();
        let int_text = if int_text.is_empty() { "0".to_string() } else { int_text };
        format!("{minus}{int_text}{gap}{}", " ".repeat(width))
    } else {
        let num = pad_fraction(n.to_string(), num_pattern, true);
        let den_text = match fixed_den {
            Some(d) => d.to_string(),
            None => pad_fraction(d.to_string(), den, false),
        };
        if int_text.is_empty() {
            format!("{minus}{num}/{den_text}")
        } else {
            format!("{minus}{int_text}{gap}{num}/{den_text}")
        }
    };
    let text = format!("{lead}{text}");
    let mut parser = Parser::new(&format!("{prefix}0"));
    parser.parse();
    let color = match parser.parts.first() {
        Some(ParsePart::Number(p)) => p.color,
        _ => None,
    };
    Some(Formatted { text, color, error: None })
}

pub fn format_number(value_original: f64, format: &str, locale: &Locale) -> Formatted {
    if let Some(f) = format_fraction(value_original, format) {
        return f;
    }
    let mut parser = Parser::new(format);
    parser.parse();
    let parts = parser.parts;
    // There are four parts:
    // 1) Positive numbers
    // 2) Negative numbers
    // 3) Zero
    // 4) Text
    // If you specify only one section of format code, the code in that section is used for all numbers.
    // If you specify two sections of format code, the first section of code is used
    // for positive numbers and zeros, and the second section of code is used for negative numbers.
    // When you skip code sections in your number format,
    // you must include a semicolon for each of the missing sections of code.
    // You can use the ampersand (&) text operator to join, or concatenate, two values.
    let mut value = value_original;
    let part;
    match parts.len() {
        1 => {
            part = &parts[0];
        }
        2 => {
            if value >= 0.0 {
                part = &parts[0]
            } else {
                value = -value;
                part = &parts[1];
            }
        }
        3 => {
            if value > 0.0 {
                part = &parts[0]
            } else if value < 0.0 {
                value = -value;
                part = &parts[1];
            } else {
                value = 0.0;
                part = &parts[2];
            }
        }
        4 => {
            if value > 0.0 {
                part = &parts[0]
            } else if value < 0.0 {
                value = -value;
                part = &parts[1];
            } else {
                value = 0.0;
                part = &parts[2];
            }
        }
        _ => {
            return Formatted {
                text: "#VALUE!".to_owned(),
                color: None,
                error: Some("Too many parts".to_owned()),
            };
        }
    }
    match part {
        ParsePart::Error(..) => Formatted {
            text: "#VALUE!".to_owned(),
            color: None,
            error: Some("Problem parsing format string".to_owned()),
        },
        ParsePart::General(..) => {
            // FIXME: This is "General formatting"
            // We should have different codepaths for general formatting and errors
            // ENS-Office: like Excel, General never shows more than 15
            // significant digits (6366933.539999999 is shown as 6366933.54).
            let mut value = crate::number_format::to_excel_precision(value, 15);
            let value_abs = value.abs();
            if (1.0e-8..1.0e+11).contains(&value_abs) {
                let mut text = format!("{value:.9}");
                text = text.trim_end_matches('0').trim_end_matches('.').to_string();
                if locale.numbers.symbols.decimal != "." {
                    text = text.replace('.', &locale.numbers.symbols.decimal.to_string());
                }
                Formatted {
                    text,
                    color: None,
                    error: None,
                }
            } else {
                if value_abs == 0.0 {
                    return Formatted {
                        text: "0".to_string(),
                        color: None,
                        error: None,
                    };
                }
                let exponent = value_abs.log10().floor();
                value /= 10.0_f64.powf(exponent);
                let sign = if exponent < 0.0 { '-' } else { '+' };
                let s = format!("{value:.5}");
                let mut text = format!(
                    "{}E{}{:02}",
                    s.trim_end_matches('0').trim_end_matches('.'),
                    sign,
                    exponent.abs()
                );
                if locale.numbers.symbols.decimal != "." {
                    text = text.replace('.', &locale.numbers.symbols.decimal.to_string());
                }
                Formatted {
                    text,
                    color: None,
                    error: None,
                }
            }
        }
        ParsePart::Date(p) => {
            let tokens = &p.tokens;
            let mut text = "".to_string();
            let time_fract = value.fract();
            let hours = (time_fract * 24.0).floor();
            let minutes = ((time_fract * 24.0 - hours) * 60.0).floor();
            let seconds = ((((time_fract * 24.0 - hours) * 60.0) - minutes) * 60.0).round();
            let date = from_excel_date(value as i64).ok();
            // ENS-Office: in Russian a month after the day number is in the
            // genitive ("15 марта 2023"), alone it is the nominative ("Март 2023").
            let with_day = tokens.iter().any(|t| matches!(t, TextToken::Day | TextToken::DayPadded));
            for token in tokens {
                match token {
                    TextToken::MonthName if with_day && genitive_month(locale, 1).is_some() => {
                        let Some(d) = date else {
                            return Formatted {
                                text: "#VALUE!".to_owned(),
                                color: None,
                                error: Some(format!("Invalid date value: '{value}'")),
                            };
                        };
                        let name = genitive_month(locale, d.month() as usize).unwrap_or_default();
                        text = format!("{text}{name}");
                    }
                    TextToken::Literal(c) => {
                        text = format!("{text}{c}");
                    }
                    TextToken::Text(t) => {
                        text = format!("{text}{t}");
                    }
                    TextToken::Ghost(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{text} ");
                    }
                    TextToken::Spacer(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{text} ");
                    }
                    TextToken::Raw => {
                        text = format!("{text}{value}");
                    }
                    TextToken::Digit(_) => {}
                    TextToken::Period => {}
                    TextToken::Day => match date {
                        Some(date) => {
                            let day = date.day() as usize;
                            text = format!("{text}{day}");
                        }
                        None => {
                            return Formatted {
                                text: "#VALUE!".to_owned(),
                                color: None,
                                error: Some(format!("Invalid date value: '{value}'")),
                            }
                        }
                    },
                    TextToken::DayPadded => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let day = date.day() as usize;
                        text = format!("{text}{day:02}");
                    }
                    TextToken::DayNameShort => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let mut day = date.weekday().number_from_monday() as usize;
                        if day == 7 {
                            day = 0;
                        }
                        text = format!("{}{}", text, locale.dates.day_names_short[day]);
                    }
                    TextToken::DayName => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let mut day = date.weekday().number_from_monday() as usize;
                        if day == 7 {
                            day = 0;
                        }
                        text = format!("{}{}", text, locale.dates.day_names[day]);
                    }
                    TextToken::Month => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let month = date.month() as usize;
                        text = format!("{text}{month}");
                    }
                    TextToken::MonthPadded => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let month = date.month() as usize;
                        text = format!("{text}{month:02}");
                    }
                    TextToken::MonthNameShort => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let month = date.month() as usize;
                        text = format!("{}{}", text, locale.dates.months_short[month - 1]);
                    }
                    TextToken::MonthName => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let month = date.month() as usize;
                        text = format!("{}{}", text, locale.dates.months[month - 1]);
                    }
                    TextToken::MonthLetter => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        let month = date.month() as usize;
                        let months_letter = &locale.dates.months_letter[month - 1];
                        text = format!("{text}{months_letter}");
                    }
                    TextToken::YearShort => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        text = format!("{}{}", text, date.format("%y"));
                    }
                    TextToken::Year => {
                        let date = match date {
                            Some(d) => d,
                            None => {
                                return Formatted {
                                    text: "#VALUE!".to_owned(),
                                    color: None,
                                    error: Some(format!("Invalid date value: '{value}'")),
                                }
                            }
                        };
                        text = format!("{}{}", text, date.year());
                    }
                    TextToken::Hour => {
                        let mut hour = hours as i32;
                        if p.use_ampm {
                            if hour == 0 {
                                hour = 12;
                            } else if hour > 12 {
                                hour -= 12;
                            }
                        }
                        text = format!("{text}{hour}");
                    }
                    TextToken::HourPadded => {
                        let mut hour = hours as i32;
                        if p.use_ampm {
                            if hour == 0 {
                                hour = 12;
                            } else if hour > 12 {
                                hour -= 12;
                            }
                        }
                        text = format!("{text}{hour:02}");
                    }
                    TextToken::Second => {
                        let second = seconds as i32;
                        text = format!("{text}{second}");
                    }
                    TextToken::SecondPadded => {
                        let second = seconds as i32;
                        text = format!("{text}{second:02}");
                    }
                    TextToken::AMPM => {
                        let ampm = if hours < 12.0 { "AM" } else { "PM" };
                        text = format!("{text}{ampm}");
                    }
                    TextToken::Minute => {
                        let minute = minutes as i32;
                        text = format!("{text}{minute}");
                    }
                    TextToken::MinutePadded => {
                        let minute = minutes as i32;
                        text = format!("{text}{minute:02}");
                    }
                    TextToken::ElapsedHour => {
                        let hour = (value * 24.0).floor() as i32;
                        text = format!("{text}{hour}");
                    }
                    TextToken::ElapsedHourPadded => {
                        let hour = (value * 24.0).floor() as i32;
                        text = format!("{text}{hour:02}");
                    }
                    TextToken::ElapsedMinute => {
                        let minute = (value * 24.0 * 60.0).floor() as i32;
                        text = format!("{text}{minute}");
                    }
                    TextToken::ElapsedMinutePadded => {
                        let minute = (value * 24.0 * 60.0).floor() as i32;
                        text = format!("{text}{minute:02}");
                    }
                    TextToken::ElapsedSecond => {
                        let second = (value * 24.0 * 60.0 * 60.0).floor() as i32;
                        text = format!("{text}{second}");
                    }
                    TextToken::ElapsedSecondPadded => {
                        let second = (value * 24.0 * 60.0 * 60.0).floor() as i32;
                        text = format!("{text}{second:02}");
                    }
                }
            }
            Formatted {
                text,
                color: p.color,
                error: None,
            }
        }
        ParsePart::Number(p) => {
            let mut text = "".to_string();
            if let Some(c) = p.currency {
                text = format!("{c}");
            }
            let tokens = &p.tokens;
            value = value * 100.0_f64.powi(p.percent) / (1000.0_f64.powi(p.comma));
            // p.precision is the number of significant digits _after_ the decimal point
            // ENS-Office: Excel rounds halves away from zero on the decimal
            // digits (1234,5 → "1235"; 2,345 → "2,35"); Rust's formatting
            // rounds exact binary ties to even.
            value = if p.is_scientific {
                // in scientific notation the precision is that of the mantissa
                to_precision(value, (p.precision as usize) + format!("{}", value.abs().floor()).len())
            } else {
                round_half_away(value, p.precision)
            };
            let mut value_abs = value.abs();
            let mut exponent_part: Vec<char> = vec![];
            let mut exponent_is_negative = value_abs < 10.0;
            if p.is_scientific {
                if value_abs == 0.0 {
                    exponent_part = vec!['0'];
                    exponent_is_negative = false;
                } else {
                    // TODO: Implement engineering formatting.
                    let exponent = value_abs.log10().floor();
                    exponent_part = format!("{}", exponent.abs()).chars().collect();
                    value /= 10.0_f64.powf(exponent);
                    value = to_precision(value, 15);
                    value_abs = value.abs();
                }
            }
            let l_exp = exponent_part.len() as i32;
            let int_number = if p.precision == 0 {
                value_abs.round()
            } else {
                value_abs.floor()
            };
            let mut int_part: Vec<char> = format!("{}", int_number).chars().collect();
            if int_number as i64 == 0 {
                int_part = vec![];
            }
            let fract_part = get_fract_part(value_abs, p.precision, int_part.len());
            // ln is the number of digits of the integer part of the value
            let ln = int_part.len() as i32;
            // digit count is the number of digit tokens ('0', '?' and '#') to the left of the decimal point
            let digit_count = p.digit_count;
            // digit_index points to the digit index in value that we have already formatted
            let mut digit_index = 0;

            let symbols = &locale.numbers.symbols;
            let group_sizes = locale.numbers.decimal_formats.standard.to_owned();
            let group_separator = symbols.group.to_owned();
            let decimal_separator = symbols.decimal.to_owned();
            // There probably are better ways to check if a number at a given precision is negative :/
            let is_negative = value < -(10.0_f64.powf(-(p.precision as f64)));
            let mut needs_period = false;

            for token in tokens {
                match token {
                    TextToken::Literal(c) => {
                        text = format!("{text}{c}");
                    }
                    TextToken::Text(t) => {
                        text = format!("{text}{t}");
                    }
                    TextToken::Ghost(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{text} ");
                    }
                    TextToken::Spacer(_) => {
                        // we just leave a whitespace
                        // This is what the TEXT function does
                        text = format!("{text} ");
                    }
                    TextToken::Raw => {
                        text = format!("{text}{value}");
                    }
                    TextToken::Period => {
                        // if !fract_part.is_empty() &&  {
                        //     text = format!("{text}{decimal_separator}");
                        // }
                        needs_period = true;
                    }
                    TextToken::Digit(digit) => {
                        if digit.number.is_integer() {
                            // 1. Integer part
                            let index = digit.index;
                            let number_index = ln - digit_count + index;
                            if index == 0 && is_negative {
                                text = format!("-{text}");
                            }
                            if ln <= digit_count {
                                // The number of digits is less or equal than the number of digit tokens
                                // i.e. the value is 123 and the format_code is ##### (ln = 3 and digit_count = 5)
                                if !(number_index < 0 && digit.kind == '#') {
                                    let c = if number_index < 0 {
                                        if digit.kind == '0' {
                                            '0'
                                        } else {
                                            // digit.kind = '?'
                                            ' '
                                        }
                                    } else {
                                        int_part[number_index as usize]
                                    };
                                    let sep = if use_group_separator(
                                        p.use_thousands,
                                        ln - digit_index,
                                        &group_sizes,
                                    ) {
                                        &group_separator
                                    } else {
                                        ""
                                    };
                                    text = format!("{text}{c}{sep}");
                                }
                                digit_index += 1;
                            } else {
                                // The number is larger than the formatting code 12345 and 0##
                                // We just hit the first formatting digit (0 in the example above) so we write as many digits as we can (123 in the example)
                                for i in digit_index..number_index + 1 {
                                    let sep = if use_group_separator(
                                        p.use_thousands,
                                        ln - i,
                                        &group_sizes,
                                    ) {
                                        &group_separator
                                    } else {
                                        ""
                                    };
                                    text = format!("{}{}{}", text, int_part[i as usize], sep);
                                }
                                digit_index = number_index + 1;
                            }
                        } else if digit.number.is_decimal() {
                            // 2. After the decimal point
                            let index = digit.index as usize;
                            if index < fract_part.len() {
                                if needs_period {
                                    text = format!(
                                        "{}{}{}",
                                        text, decimal_separator, fract_part[index]
                                    );
                                } else {
                                    text = format!("{}{}", text, fract_part[index]);
                                }
                            } else if digit.kind == '0' {
                                if needs_period {
                                    text = format!("{text}{}0", decimal_separator);
                                } else {
                                    text = format!("{text}0");
                                }
                            } else if digit.kind == '?' {
                                text = format!("{text} ");
                            } else if digit.kind == '#' && needs_period {
                                // FIXME: This is what Excel does, but it transforms "3" into "3."
                                text = format!("{text}{}", decimal_separator);
                            }
                            needs_period = false;
                        } else if digit.number.is_exponent() {
                            // 3. Exponent part
                            let index = digit.index;
                            if index == 0 {
                                if exponent_is_negative {
                                    text = format!("{text}E-");
                                } else if p.scientific_minus {
                                    text = format!("{text}E");
                                } else {
                                    text = format!("{text}E+");
                                }
                            }
                            let number_index = l_exp - (p.exponent_digit_count - index);
                            if l_exp <= p.exponent_digit_count {
                                if !(number_index < 0 && digit.kind == '#') {
                                    let c = if number_index < 0 {
                                        if digit.kind == '?' {
                                            ' '
                                        } else {
                                            '0'
                                        }
                                    } else {
                                        exponent_part[number_index as usize]
                                    };

                                    text = format!("{text}{c}");
                                }
                            } else {
                                for i in 0..number_index + 1 {
                                    text = format!("{}{}", text, exponent_part[i as usize]);
                                }
                                digit_index += number_index + 1;
                            }
                        }
                    }
                    // Date tokens should not be present
                    TextToken::Day => {}
                    TextToken::DayPadded => {}
                    TextToken::DayNameShort => {}
                    TextToken::DayName => {}
                    TextToken::Month => {}
                    TextToken::MonthPadded => {}
                    TextToken::MonthNameShort => {}
                    TextToken::MonthName => {}
                    TextToken::MonthLetter => {}
                    TextToken::YearShort => {}
                    TextToken::Year => {}
                    TextToken::Hour => {}
                    TextToken::HourPadded => {}
                    TextToken::Minute => {}
                    TextToken::MinutePadded => {}
                    TextToken::Second => {}
                    TextToken::SecondPadded => {}
                    TextToken::AMPM => {}
                    TextToken::ElapsedHour => {}
                    TextToken::ElapsedHourPadded => {}
                    TextToken::ElapsedMinute => {}
                    TextToken::ElapsedMinutePadded => {}
                    TextToken::ElapsedSecond => {}
                    TextToken::ElapsedSecondPadded => {}
                }
            }
            Formatted {
                text,
                color: p.color,
                error: None,
            }
        }
    }
}

fn parse_day(day_str: &str) -> Result<(u32, String), String> {
    let bytes = day_str.bytes();
    let bytes_len = bytes.len();
    if bytes_len <= 2 {
        match day_str.parse::<u32>() {
            Ok(y) => {
                if bytes_len == 2 {
                    return Ok((y, "dd".to_string()));
                } else {
                    return Ok((y, "d".to_string()));
                }
            }
            Err(_) => return Err("Not a valid day".to_string()),
        }
    }
    Err("Not a valid day".to_string())
}

fn parse_month(month_str: &str, locale: &Locale) -> Result<(u32, String), String> {
    let bytes = month_str.bytes();
    let bytes_len = bytes.len();
    if bytes_len <= 2 {
        match month_str.parse::<u32>() {
            Ok(y) => {
                if bytes_len == 2 {
                    return Ok((y, "mm".to_string()));
                } else {
                    return Ok((y, "m".to_string()));
                }
            }
            Err(_) => return Err("Not a valid month".to_string()),
        }
    }

    let month_names_short = &locale.dates.months_short;
    let month_names_long = &locale.dates.months;
    if let Some(m) = month_names_short.iter().position(|r| r == month_str) {
        return Ok((m as u32 + 1, "mmm".to_string()));
    }
    if let Some(m) = month_names_long.iter().position(|r| r == month_str) {
        return Ok((m as u32 + 1, "mmmm".to_string()));
    }
    Err("Not a valid month".to_string())
}

fn parse_year(year_str: &str) -> Result<(i32, String), String> {
    // year is either 2 digits or 4 digits
    // 23 -> 2023
    // 75 -> 1975
    // 30 is the split number (yeah, that's not going to be a problem any time soon)
    // 30 => 1930
    // 29 => 2029
    let bytes = year_str.bytes();
    let bytes_len = bytes.len();
    if bytes_len != 2 && bytes_len != 4 {
        return Err("Not a valid year".to_string());
    }
    match year_str.parse::<i32>() {
        Ok(y) => {
            if y < 30 {
                Ok((2000 + y, "yy".to_string()))
            } else if y < 100 {
                Ok((1900 + y, "yy".to_string()))
            } else {
                Ok((y, "yyyy".to_string()))
            }
        }
        Err(_) => Err("Not a valid year".to_string()),
    }
}

// Check if it is a date. Other spreadsheet engines support a wide variety of dates formats
// Here we support a small subset of them.
//
// The grammar is:
//
// date -> long_date | short_date | iso-date
// short_date -> month separator year
// long_date -> day separator month separator year
// iso_date -> long_year separator number_month separator number_day
// separator -> "/" | "-" | "."
// day -> number | padded number
// month -> number_month | name_month
// number_month -> number | padded number |
// name_month -> short name | full name
// year -> short_year | long year
//
// NOTE 1: The separator has to be the same
// NOTE 2: In some engines "2/3" is implemented ad "2/March of the present year"
// NOTE 3: I did not implement the "short date"
pub(crate) fn parse_date(value: &str, locale: &Locale) -> Result<(i32, String), String> {
    let separator = if value.contains('/') {
        '/'
    } else if value.contains('-') {
        '-'
    } else if value.contains('.') {
        '.'
    } else {
        return Err("Not a valid date".to_string());
    };

    let parts: Vec<&str> = value.split(separator).collect();
    let mut is_iso_date = false;
    let mut day_first = true;
    let (day_str, month_str, year_str) = if parts.len() == 3 {
        if parts[0].len() == 4 {
            // ISO date  yyyy-mm-dd
            if !parts[1].chars().all(char::is_numeric) {
                return Err("Not a valid date".to_string());
            }
            if !parts[2].chars().all(char::is_numeric) {
                return Err("Not a valid date".to_string());
            }
            is_iso_date = true;
            (parts[2], parts[1], parts[0])
        } else {
            //  localized date dd-mm-yyyy or mm-dd-yyyy
            // TODO: A bit hacky, but works for now
            if locale.dates.date_formats.short.starts_with('d') {
                (parts[0], parts[1], parts[2])
            } else {
                day_first = false;
                (parts[1], parts[0], parts[2])
            }
        }
    } else {
        return Err("Not a valid date".to_string());
    };
    let (day, day_format) = parse_day(day_str)?;
    let (month, month_format) = parse_month(month_str, locale)?;
    let (year, year_format) = parse_year(year_str)?;
    let serial_number = match date_to_serial_number(day, month, year) {
        Ok(n) => n,
        Err(_) => return Err("Not a valid date".to_string()),
    };
    if is_iso_date {
        Ok((
            serial_number,
            format!("yyyy{separator}{month_format}{separator}{day_format}"),
        ))
    } else if !day_first {
        Ok((
            serial_number,
            format!("{month_format}{separator}{day_format}{separator}{year_format}"),
        ))
    } else {
        Ok((
            serial_number,
            format!("{day_format}{separator}{month_format}{separator}{year_format}"),
        ))
    }
}

/// Parses a formatted number, returning the numeric value together with the format
/// Uses heuristics to guess the format string
/// "$ 123,345.678" => (123345.678, "$#,##0.00")
/// "30.34%" => (0.3034, "0.00%")
/// 100€ => (100, "100€")
pub(crate) fn parse_formatted_number(
    original: &str,
    currencies: &[&str],
    locale: &Locale,
) -> Result<(f64, Option<String>), String> {
    let value = original.trim();
    let scientific_format = "0.00E+00";

    let symbols = &locale.numbers.symbols;
    let decimal_separator = symbols.decimal.chars().next().unwrap_or('.');
    let group_separator = symbols.group.chars().next().unwrap_or(',');

    // Check if it is a percentage
    if let Some(p) = value.strip_suffix('%') {
        let (f, options) = parse_number(p.trim(), decimal_separator, group_separator)?;
        if options.is_scientific {
            return Ok((f / 100.0, Some(scientific_format.to_string())));
        }
        // We ignore the separator
        if options.decimal_digits > 0 {
            // Percentage format with decimals
            return Ok((f / 100.0, Some("#,##0.00%".to_string())));
        }
        // Percentage format standard
        return Ok((f / 100.0, Some("#,##0%".to_string())));
    }

    // check if it is a currency in currencies
    for currency in currencies {
        if let Some(p) = value.strip_prefix(&format!("-{currency}")) {
            let (f, options) = parse_number(p.trim(), decimal_separator, group_separator)?;
            if options.is_scientific {
                return Ok((f, Some(scientific_format.to_string())));
            }
            if options.decimal_digits > 0 {
                return Ok((-f, Some(format!("{currency}#,##0.00"))));
            }
            return Ok((-f, Some(format!("{currency}#,##0"))));
        } else if let Some(p) = value.strip_prefix(currency) {
            let (f, options) = parse_number(p.trim(), decimal_separator, group_separator)?;
            if options.is_scientific {
                return Ok((f, Some(scientific_format.to_string())));
            }
            if options.decimal_digits > 0 {
                return Ok((f, Some(format!("{currency}#,##0.00"))));
            }
            return Ok((f, Some(format!("{currency}#,##0"))));
        } else if let Some(p) = value.strip_suffix(currency) {
            let (f, options) = parse_number(p.trim(), decimal_separator, group_separator)?;
            if options.is_scientific {
                return Ok((f, Some(scientific_format.to_string())));
            }
            if options.decimal_digits > 0 {
                let currency_format = &format!("#,##0.00{currency}");
                return Ok((f, Some(currency_format.to_string())));
            }
            let currency_format = &format!("#,##0{currency}");
            return Ok((f, Some(currency_format.to_string())));
        }
    }

    // check if it is a date. NOTE: we don't trim the original here
    if let Ok((serial_number, format)) = parse_date(original, locale) {
        return Ok((serial_number as f64, Some(format)));
    }

    // Lastly we check if it is a number
    let (f, options) = parse_number(value, decimal_separator, group_separator)?;
    if options.is_scientific {
        return Ok((f, Some(scientific_format.to_string())));
    }
    if options.has_commas {
        if options.decimal_digits > 0 {
            // group separator and two decimal points
            return Ok((f, Some("#,##0.00".to_string())));
        }
        // Group separator and no decimal points
        return Ok((f, Some("#,##0".to_string())));
    }
    Ok((f, None))
}

struct NumberOptions {
    has_commas: bool,
    is_scientific: bool,
    decimal_digits: usize,
}

// tries to parse 'value' as a number.
// If it is a number it either uses commas as thousands separator or it does not
fn parse_number(
    value: &str,
    decimal_separator: char,
    group_separator: char,
) -> Result<(f64, NumberOptions), String> {
    let mut position = 0;
    let characters: Vec<char> = value.chars().collect();
    let len = characters.len();
    if len == 0 {
        return Err("Cannot parse number".to_string());
    }
    let mut chars = String::from("");
    let mut group_separator_index = Vec::new();
    // ENS-Office: where the group separator is a (non-breaking) space, as in
    // Russian, people type and paste ordinary spaces: "1 234 567,89".
    let spaced = group_separator.is_whitespace();
    let is_group = |x: char| x == group_separator || (spaced && (x == ' ' || x == '\u{a0}' || x == '\u{202f}'));
    // get the sign
    let sign = if characters[0] == '-' {
        position += 1;
        -1.0
    } else if characters[0] == '+' {
        position += 1;
        1.0
    } else {
        1.0
    };

    if position >= len {
        return Err("Cannot parse number".to_string());
    }

    if is_group(characters[position]) {
        return Err("Cannot parse number".to_string());
    }

    // numbers before the decimal point
    while position < len {
        let x = characters[position];
        if x.is_ascii_digit() {
            chars.push(x);
        } else if is_group(x) {
            group_separator_index.push(chars.len());
        } else {
            break;
        }
        position += 1;
    }
    // Check the group separator is in multiples of three
    for index in &group_separator_index {
        if (chars.len() - index) % 3 != 0 {
            return Err("Cannot parse number".to_string());
        }
    }
    let mut decimal_digits = 0;
    if position < len && characters[position] == decimal_separator {
        // numbers after the decimal point
        chars.push('.');
        position += 1;
        let start_position = 0;
        while position < len {
            let x = characters[position];
            if x.is_ascii_digit() {
                chars.push(x);
            } else {
                break;
            }
            position += 1;
        }
        decimal_digits = position - start_position;
    }
    let mut is_scientific = false;
    if position + 1 < len && (characters[position] == 'e' || characters[position] == 'E') {
        // exponential side
        is_scientific = true;
        let x = characters[position + 1];
        if x == '-' || x == '+' || x.is_ascii_digit() {
            chars.push('e');
            chars.push(x);
            position += 2;
            while position < len {
                let x = characters[position];
                if x.is_ascii_digit() {
                    chars.push(x);
                } else {
                    break;
                }
                position += 1;
            }
        }
    }
    if position != len {
        return Err("Could not parse number".to_string());
    };
    match parse_finite_number(&chars) {
        None => Err("Failed to parse to double".to_string()),
        Some(v) => Ok((
            sign * v,
            NumberOptions {
                has_commas: !group_separator_index.is_empty(),
                is_scientific,
                decimal_digits,
            },
        )),
    }
}
