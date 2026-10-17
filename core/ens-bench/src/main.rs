//! ENS-Office: замер движка таблиц на том же синтетическом листе, что и в
//! prototypes/grid-bench (docs/research/proto-01-grid-bench.md).
//!
//! Запуск: cargo run --release -- 50000 [--no-vlookup]

use ironcalc_base::UserModel;
use std::time::Instant;

const LOOKUP_ROWS: usize = 2000;

/// mulberry32 — тот же генератор, что в data.js, чтобы данные совпадали.
struct Rng(u32);
impl Rng {
    fn next(&mut self) -> f64 {
        self.0 = self.0.wrapping_add(0x6d2b_79f5);
        let mut t = self.0;
        t = (t ^ (t >> 15)).wrapping_mul(1 | t);
        t = (t.wrapping_add((t ^ (t >> 7)).wrapping_mul(61 | t))) ^ t;
        ((t ^ (t >> 14)) as f64) / 4_294_967_296.0
    }
}

fn ms(t: Instant) -> u128 {
    t.elapsed().as_millis()
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let rows: usize = args.get(1).and_then(|s| s.parse().ok()).unwrap_or(50_000);
    let use_lookup = !args.iter().any(|a| a == "--no-vlookup");

    let mut rng = Rng(42);
    let lookup: Vec<(String, f64)> = (0..LOOKUP_ROWS)
        .map(|i| (format!("К{i:04}"), (rng.next() * 2000.0).round() / 100.0))
        .collect();

    let t = Instant::now();
    let mut m = UserModel::new_empty("Bench", "en", "UTC", "en").expect("model");
    m.rename_sheet(0, "Data").expect("rename");
    m.new_sheet().expect("sheet");
    m.rename_sheet(1, "Lookup").expect("rename");
    m.pause_evaluation();
    for (i, (k, v)) in lookup.iter().enumerate() {
        let r = i as i32 + 1;
        m.set_user_input(1, r, 1, k).expect("set");
        m.set_user_input(1, r, 2, &v.to_string()).expect("set");
    }
    let headers = ["№", "Ключ", "Филиал", "Кол-во", "Цена", "Ставка", "Сумма", "Размер", "Итог", "ВСЕГО"];
    for (c, h) in headers.iter().enumerate() {
        m.set_user_input(0, 1, c as i32 + 1, h).expect("set");
    }
    let mut first_key_row = 0;
    for i in 0..rows {
        let n = i + 2;
        let li = (rng.next() * LOOKUP_ROWS as f64).floor() as usize;
        if i == 0 {
            first_key_row = li + 1;
        }
        let qty = 1 + (rng.next() * 50.0).floor() as i64;
        let price = (rng.next() * 10000.0).round() / 100.0;
        let r = n as i32;
        let branch = ["Д", "Л", "З", "Х"][i % 4];
        let lookup_formula = if use_lookup {
            format!("=VLOOKUP(B{n},Lookup!$A$1:$B${LOOKUP_ROWS},2,0)")
        } else {
            format!("=Lookup!B{}", li + 1)
        };
        let cells = [
            (i + 1).to_string(),
            lookup[li].0.clone(),
            branch.to_string(),
            qty.to_string(),
            price.to_string(),
            lookup_formula,
            format!("=D{n}*E{n}"),
            format!("=IF(G{n}>500,\"крупн\",\"мелк\")"),
            format!("=ROUND(G{n}*F{n}/100,2)"),
        ];
        for (c, v) in cells.iter().enumerate() {
            m.set_user_input(0, r, c as i32 + 1, v).expect("set");
        }
        if i == 0 {
            m.set_user_input(0, r, 10, &format!("=SUM(I2:I{})", rows + 1)).expect("set");
        }
    }
    let load = ms(t);

    let t = Instant::now();
    m.resume_evaluation();
    m.evaluate();
    let calc = ms(t);
    let total = m.get_formatted_cell_value(0, 2, 10).expect("total");
    eprintln!("first_calc={calc}ms load={load}ms");

    // Правка ставки в справочнике (затрагивает ~rows/2000 строк и итог).
    // set_user_input сам пересчитывает (частично, если возможно).
    let t = Instant::now();
    m.set_user_input(1, first_key_row as i32, 2, "99.99").expect("edit");
    let edit = t.elapsed().as_secs_f64() * 1000.0;
    let total2 = m.get_formatted_cell_value(0, 2, 10).expect("total");

    // Правка цены в одной строке данных (одна строка и итог).
    let t = Instant::now();
    m.set_user_input(0, (rows / 2 + 2) as i32, 5, "123.45").expect("edit");
    let edit_row = t.elapsed().as_secs_f64() * 1000.0;
    let total3 = m.get_formatted_cell_value(0, 2, 10).expect("total");

    // Контроль: полный пересчёт той же модели даёт тот же итог.
    let t = Instant::now();
    m.evaluate_full();
    let full = ms(t);
    let total_full = m.get_formatted_cell_value(0, 2, 10).expect("total");
    assert_eq!(total3, total_full, "incremental and full totals differ");

    println!(
        "rows={rows} vlookup={use_lookup} load={load}ms first_calc={calc}ms          edit_rate={edit:.1}ms edit_price={edit_row:.1}ms full_recalc={full}ms          total={total} after_rate={total2} after_price={total3}"
    );
}
