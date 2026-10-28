#![allow(clippy::unwrap_used)]

//! ENS-Office: Russian function names (=СУММ), English names still accepted
//! (=SUM), formulas shown in Russian, as in a Russian Excel.

use crate::UserModel;

fn model() -> UserModel<'static> {
    UserModel::new_empty("model", "ru", "UTC", "ru").unwrap()
}

#[test]
fn ens_ru_function_names() {
    let mut m = model();
    m.set_user_input(0, 1, 1, "2").unwrap();
    m.set_user_input(0, 2, 1, "3,5").unwrap();
    m.set_user_input(0, 3, 1, "=СУММ(A1:A2)").unwrap();
    m.set_user_input(0, 4, 1, "=сумм(A1:A2)").unwrap();
    m.set_user_input(0, 5, 1, "=SUM(A1:A2)").unwrap();
    m.set_user_input(0, 6, 1, "=ЕСЛИ(A3>5;\"много\";\"мало\")").unwrap();
    m.set_user_input(0, 7, 1, "=ОКРУГЛ(СРЗНАЧ(A1:A2);1)").unwrap();
    m.set_user_input(0, 8, 1, "=ВПР(2;A1:A2;1;ЛОЖЬ)").unwrap();
    m.set_user_input(0, 9, 1, "=1/0").unwrap();
    m.set_user_input(0, 10, 1, "=ИСТИНА").unwrap();
    for row in 3..=5 {
        assert_eq!(m.get_formatted_cell_value(0, row, 1).unwrap(), "5,5", "row {row}");
    }
    // English input is shown in Russian.
    assert_eq!(m.get_cell_content(0, 5, 1).unwrap(), "=СУММ(A1:A2)");
    assert_eq!(m.get_formatted_cell_value(0, 6, 1).unwrap(), "много");
    assert_eq!(m.get_formatted_cell_value(0, 7, 1).unwrap(), "2,8");
    assert_eq!(m.get_formatted_cell_value(0, 8, 1).unwrap(), "2");
    assert_eq!(m.get_formatted_cell_value(0, 9, 1).unwrap(), "#ДЕЛ/0!");
    assert_eq!(m.get_formatted_cell_value(0, 10, 1).unwrap(), "ИСТИНА");
}

#[test]
fn ens_english_language_still_works() {
    let mut m = UserModel::new_empty("model", "en", "UTC", "en").unwrap();
    m.set_user_input(0, 1, 1, "=SUM(1,2)").unwrap();
    assert_eq!(m.get_formatted_cell_value(0, 1, 1).unwrap(), "3");
}
