/// Select sheet by id or by name.
#[derive(Clone, Debug)]
pub enum SheetSelector {
    ById(usize),
    ByName(String),
}

impl SheetSelector {
    pub fn find_in<'a>(&self, sheetnames: &'a [String]) -> Result<&'a String, String> {
        match self {
            SheetSelector::ById(id) => {
                if *id >= sheetnames.len() {
                    Err(format!(
                        "sheet id `{}` is not valid - only **{}** sheets avaliable!",
                        id,
                        sheetnames.len()
                    ))
                } else {
                    Ok(&sheetnames[*id])
                }
            }
            SheetSelector::ByName(name) => {
                if let Some(name) = sheetnames.iter().find(|s| *s == name) {
                    Ok(name)
                } else {
                    let msg = format!(
                        "sheet name `{}` is not in ({})",
                        name,
                        sheetnames.join(", ")
                    );
                    Err(msg)
                }
            }
        }
    }
}

impl std::str::FromStr for SheetSelector {
    type Err = String;
    fn from_str(str: &str) -> Result<Self, Self::Err> {
        match str.parse() {
            Ok(id) => Ok(SheetSelector::ById(id)),
            Err(_) => Ok(SheetSelector::ByName(str.to_string())),
        }
    }
}
