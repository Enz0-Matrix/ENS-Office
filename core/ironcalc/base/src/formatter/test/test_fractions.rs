#![allow(clippy::unwrap_used)]

use crate::{formatter::format::format_number, locale::get_locale};

fn f(v: f64, code: &str) -> String {
    format_number(v, code, get_locale("en").unwrap()).text
}

#[test]
fn fractions() {
    assert_eq!(f(1.5, "# ?/?"), "1 1/2");
    assert_eq!(f(0.75, "# ?/?"), "3/4");
    assert_eq!(f(3.0, "# ?/?"), "3    ");
    assert_eq!(f(-2.25, "# ?/?"), "-2 1/4");
    assert_eq!(f(0.333, "# ??/??"), " 1/3 ");
    assert_eq!(f(3.14159, "# ??/??"), "3 14/99");
    assert_eq!(f(0.3125, "# ?/8"), "3/8");
    assert_eq!(f(1.37, "# ??/100"), "1 37/100");
    assert_eq!(f(1.5, "?/?"), "3/2");
    assert_eq!(f(0.999, "# ?/?"), "1    ");
    let r = format_number(-0.5, "# ?/?;[Red]-# ?/?", get_locale("en").unwrap());
    assert_eq!(r.text, "-1/2");
    assert!(r.color.is_some());
    // dates with '/' are not fractions
    assert_eq!(f(45000.0, "dd/mm/yyyy"), "15/03/2023");
}
