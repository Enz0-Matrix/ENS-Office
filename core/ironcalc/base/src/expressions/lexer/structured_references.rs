// Grammar:
// structured references -> table_name "[" arguments "]"
//  arguments -> table_reference | "["specifier"]"  "," table_reference
//  specifier > "#All"      |
//              "#This Row" |
//              "#Data"     |
//              "#Headers"  |
//              "#Totals"
// table_reference -> column_reference | range_reference
// column reference -> column_name | "["column_name"]"
// range_reference -> column_reference":"column_reference

use crate::expressions::token::TokenType;
use crate::expressions::token::{TableReference, TableSpecifier};

use super::Result;
use super::{Lexer, LexerError};

impl<'a> Lexer<'a> {
    fn consume_table_specifier(&mut self) -> Result<Option<TableSpecifier>> {
        if self.peek_char() == Some('#') {
            // ENS-Office: как в английском и русском Excel, без учёта регистра (ADR-013)
            const NAMES: [(&str, TableSpecifier); 10] = [
                ("#this row]", TableSpecifier::ThisRow),
                ("#all]", TableSpecifier::All),
                ("#data]", TableSpecifier::Data),
                ("#headers]", TableSpecifier::Headers),
                ("#totals]", TableSpecifier::Totals),
                ("#эта строка]", TableSpecifier::ThisRow),
                ("#все]", TableSpecifier::All),
                ("#данные]", TableSpecifier::Data),
                ("#заголовки]", TableSpecifier::Headers),
                ("#итоги]", TableSpecifier::Totals),
            ];
            let rest: String = self.chars[self.position..self.len].iter().collect::<String>().to_lowercase();
            for (name, specifier) in NAMES {
                if rest.starts_with(name) {
                    self.position += name.chars().count();
                    return Ok(Some(specifier));
                }
            }
            Err(LexerError {
                position: self.position,
                message: "Invalid structured reference".to_string(),
            })
        } else {
            Ok(None)
        }
    }

    /// ENS-Office: имя столбца без скобок после `@` (`[@Цена]`) — до `]`; `'` экранирует знак.
    fn consume_bare_column(&mut self) -> Result<String> {
        let mut out = String::new();
        while self.position < self.len {
            let ch = self.chars[self.position];
            if ch == ']' {
                return Ok(out.trim().to_string());
            }
            if ch == '\'' && self.position + 1 < self.len {
                self.position += 1;
                out.push(self.chars[self.position]);
            } else {
                out.push(ch);
            }
            self.position += 1;
        }
        Err(LexerError {
            position: self.position,
            message: "Invalid column name".to_string(),
        })
    }

    /// ENS-Office: после `[@`: `]`, `Цена]`, `[Цена с НДС]]`, `[Янв]:[Мар]]`.
    fn consume_this_row_short(&mut self, table_name: &str) -> Result<TokenType> {
        let reference = if self.peek_char() == Some(']') {
            None
        } else if self.peek_char() == Some('[') {
            let left = self.consume_column_reference()?;
            if self.peek_char() == Some(':') {
                self.position += 1;
                let right = self.consume_column_reference()?;
                Some(TableReference::RangeReference((left, right)))
            } else {
                Some(TableReference::ColumnReference(left))
            }
        } else {
            Some(TableReference::ColumnReference(self.consume_bare_column()?))
        };
        self.expect(TokenType::RightBracket)?;
        Ok(TokenType::StructuredReference {
            table_name: table_name.to_string(),
            specifier: Some(TableSpecifier::ThisRow),
            table_reference: reference,
        })
    }

    fn consume_column_reference(&mut self) -> Result<String> {
        self.consume_whitespace();
        let end_char = if self.peek_char() == Some('[') {
            self.position += 1;
            ']'
        } else {
            ')'
        };

        let mut position = self.position;
        while position < self.len {
            let next_char = self.chars[position];
            if next_char != end_char {
                position += 1;
                if next_char == '\'' {
                    if position == self.len {
                        return Err(LexerError {
                            position: self.position,
                            message: "Invalid column name".to_string(),
                        });
                    }
                    // skip next char
                    position += 1
                }
            } else {
                break;
            }
        }
        let chars: String = self.chars[self.position..position].iter().collect();
        if end_char == ']' {
            position += 1;
        }
        self.position = position;
        Ok(chars
            .replace("'[", "[")
            .replace("']", "]")
            .replace("'#", "#")
            .replace("'@", "@")
            .replace("''", "'"))
    }

    /// Possibilities:
    ///  1. MyTable[#Totals] or MyTable[#This Row]
    ///  2. MyTable[MyColumn]
    ///  3. MyTable[[My Column]]
    ///  4. MyTable[[#This Row], [My Column]]
    ///  5. MyTable[[#Totals], [MyColumn]]
    ///  6. MyTable[[#This Row], [Jan]:[Dec]]
    ///  7. MyTable[]
    ///
    /// Multiple specifiers are not supported yet:
    ///  1. MyTable[[#Data], [#Totals], [MyColumn]]
    ///
    /// In particular note that names of columns are escaped only when they are in the first argument
    /// We use '[' and ']'
    /// When there is only a specifier but not a reference the specifier is not in brackets
    ///
    /// Invalid:
    /// * MyTable[#Totals, [Jan]:[March]] => MyTable[[#Totals], [Jan]:[March]]
    //
    // NOTES:
    // * MyTable[[#Totals]] is translated into MyTable[#Totals]
    // * Excel shows '@' instead of '#This Row':
    //     MyTable[[#This Row], [Jan]:[Dec]] => MyTable[@[Jan]:[Dec]]
    //   But this is only a UI thing that we will ignore for now.
    pub(crate) fn consume_structured_reference(&mut self, table_name: &str) -> Result<TokenType> {
        let outer = self.in_structured;
        self.in_structured = true;
        let result = self.consume_structured_reference_inner(table_name);
        self.in_structured = outer;
        result
    }

    fn consume_structured_reference_inner(&mut self, table_name: &str) -> Result<TokenType> {
        self.expect(TokenType::LeftBracket)?;
        let peek_char = self.peek_char();
        if peek_char == Some(']') {
            // This is just a reference to the full table
            self.expect(TokenType::RightBracket)?;
            return Ok(TokenType::StructuredReference {
                table_name: table_name.to_string(),
                specifier: None,
                table_reference: None,
            });
        }
        if peek_char == Some('@') {
            // ENS-Office: Таблица1[@Цена] — то же, что Таблица1[[#This Row],[Цена]]
            self.position += 1;
            return self.consume_this_row_short(table_name);
        }
        if peek_char == Some('#') {
            // Expecting MyTable[#Totals]
            if let Some(specifier) = self.consume_table_specifier()? {
                return Ok(TokenType::StructuredReference {
                    table_name: table_name.to_string(),
                    specifier: Some(specifier),
                    table_reference: None,
                });
            } else {
                return Err(LexerError {
                    position: self.position,
                    message: "Invalid structured reference".to_string(),
                });
            }
        } else if peek_char != Some('[') {
            // Expecting MyTable[MyColumn]
            self.position -= 1;
            let column_name = self.consume_column_reference()?;
            return Ok(TokenType::StructuredReference {
                table_name: table_name.to_string(),
                specifier: None,
                table_reference: Some(TableReference::ColumnReference(column_name)),
            });
        }
        self.expect(TokenType::LeftBracket)?;
        let specifier = self.consume_table_specifier()?;
        if specifier.is_some() {
            let peek_token = self.peek_token();
            if peek_token == TokenType::Comma || peek_token == TokenType::Semicolon {
                self.advance_token();
                self.expect(TokenType::LeftBracket)?;
            } else if peek_token == TokenType::RightBracket {
                return Ok(TokenType::StructuredReference {
                    table_name: table_name.to_string(),
                    specifier,
                    table_reference: None,
                });
            }
        }

        // Now it's either:
        // [Column Name]
        // [Column Name]:[Column Name]
        self.position -= 1;
        let column_reference = self.consume_column_reference()?;
        let table_reference = if self.peek_char() == Some(':') {
            self.position += 1;
            let column_reference_right = self.consume_column_reference()?;
            self.expect(TokenType::RightBracket)?;
            Some(TableReference::RangeReference((
                column_reference,
                column_reference_right,
            )))
        } else {
            self.expect(TokenType::RightBracket)?;
            Some(TableReference::ColumnReference(column_reference))
        };
        Ok(TokenType::StructuredReference {
            table_name: table_name.to_string(),
            specifier,
            table_reference,
        })
    }
}
