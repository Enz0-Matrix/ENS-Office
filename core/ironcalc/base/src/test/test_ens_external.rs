#![allow(clippy::unwrap_used)]

//! ENS-Office: ссылки на другие книги (в том числе закрытые).

use crate::{Model, UserModel};

/// Книга-источник «Цены.xlsx»: лист «Цены» (вид топлива → цена) и «Остатки».
fn source() -> Model<'static> {
    let mut m = Model::new_empty("Цены", "ru", "UTC", "ru").unwrap();
    m.rename_sheet_by_index(0, "Цены").unwrap();
    m.add_sheet("Остатки на складе").unwrap();
    for (r, (name, price)) in [("АИ-92", "52,3"), ("АИ-95", "57,1"), ("ДТ", "66,8")]
        .iter()
        .enumerate()
    {
        m.set_user_input(0, r as i32 + 2, 1, name.to_string())
            .unwrap();
        m.set_user_input(0, r as i32 + 2, 2, price.to_string())
            .unwrap();
    }
    m.set_user_input(0, 1, 1, "Топливо".to_string()).unwrap();
    m.set_user_input(1, 1, 1, "1200".to_string()).unwrap();
    m.set_user_input(1, 2, 1, "=A1*2".to_string()).unwrap();
    m.evaluate();
    m
}

#[test]
fn ens_external_vlookup_from_closed_book() {
    let src = source();
    let mut m = UserModel::new_empty("Отчёт", "ru", "UTC", "ru").unwrap();
    m.set_user_input(0, 1, 1, "АИ-95").unwrap();
    m.set_user_input(0, 1, 2, "=ВПР(A1;[Цены.xlsx]Цены!$A$2:$B$4;2;ЛОЖЬ)")
        .unwrap();
    m.set_user_input(0, 2, 2, "=СУММ('[Цены.xlsx]Остатки на складе'!A1:A2)")
        .unwrap();
    m.set_user_input(0, 3, 2, r"='C:\Отчёты\[Цены.xlsx]Цены'!B2*10")
        .unwrap();
    // источник ещё не прочитан: ссылки есть, значений нет
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "#Н/Д");
    // формула показывается со ссылкой на книгу
    assert_eq!(
        m.get_cell_content(0, 1, 2).unwrap(),
        "=ВПР(A1;'[Цены.xlsx]Цены'!$A$2:$B$4;2;ЛОЖЬ)"
    );
    // служебные листы скрыты и стоят после настоящих
    let props = m.get_worksheets_properties();
    assert_eq!(props.len(), 3);
    assert_eq!(props[1].state, "veryHidden");
    // путь из формулы запомнен
    assert_eq!(
        m.get_model().workbook.ens_external_books[0].target,
        r"C:\Отчёты\Цены.xlsx"
    );

    // «Обновить связи»: книга-источник прочитана с диска (здесь — модель в памяти)
    let missing = m.update_external_book("Цены.xlsx", &src).unwrap();
    assert!(missing.is_empty());
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "57,1");
    assert_eq!(m.get_formatted_cell_value(0, 2, 2).unwrap(), "3600");
    assert_eq!(m.get_formatted_cell_value(0, 3, 2).unwrap(), "523");

    // новый лист пользователя встаёт перед служебными
    m.new_sheet().unwrap();
    let props = m.get_worksheets_properties();
    assert_eq!(props[1].state, "visible");
    assert_eq!(m.get_formatted_cell_value(0, 1, 2).unwrap(), "57,1");
}
