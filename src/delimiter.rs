use std::fmt;

#[derive(Clone, Copy, Debug)]
pub struct Delimiter(pub u8);

/// Delimiter represents values that can be passed from the command line that
/// can be used as a field delimiter in CSV data.
///
/// Its purpose is to ensure that the Unicode character given decodes to a
/// valid ASCII character as required by the CSV parser.
impl Delimiter {
    pub fn as_byte(&self) -> u8 {
        self.0
    }

    pub fn as_char(&self) -> char {
        self.0 as char
    }

    pub fn to_file_extension(self) -> String {
        match self.0 {
            b'\t' => "tsv".into(),
            _ => "csv".to_string(),
        }
    }
}

impl fmt::Display for Delimiter {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.as_char())
    }
}

impl std::str::FromStr for Delimiter {
    type Err = String;
    fn from_str(str: &str) -> Result<Delimiter, Self::Err> {
        match str {
            r"\t" => Ok(Delimiter(b'\t')),
            r"\n" => Ok(Delimiter(b'\n')),
            s => {
                if s.len() != 1 {
                    let msg = format!("Could not convert '{}' to a single ASCII character.", s);
                    return Err(msg);
                }
                let c = s.chars().next().unwrap();
                if c.is_ascii() {
                    Ok(Delimiter(c as u8))
                } else {
                    let msg = format!("Could not convert '{}' to ASCII delimiter.", c);
                    Err(msg)
                }
            }
        }
    }
}
