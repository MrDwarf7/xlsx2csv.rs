use crate::delimiter::Delimiter;
use crate::sheet_selector::SheetSelector;

use clap::Parser;
use std::path::PathBuf;

/// A fast Excel-like spreadsheet to CSV coverter in Rust.
///
/// A simple usage like this:
///
/// ```
/// xlsx2csv input.xlsx sheet1.csv sheet2.csv
/// ```
///
/// If no output position args setted, it'll write first sheet to stdout.
/// So the two commands are equal:
///
/// - `xlsx2csv input.xlsx sheet1.csv`
///
/// - `xlsx2csv input.xlsx > sheet1.csv`.
///
/// If you want to select specific sheet to stdout, use `-s/--select <id or name>` (id is 0-based):
///
/// `xlsx2csv input.xlsx -s 1`
///
/// In previous command, it'll output the second sheet to stdout.
///
/// If there's many sheets that you don't wanna set filename for each,
/// use `-u` to write with sheetnames.
///
/// ```
/// xlsx2csv input.xlsx -u
/// ```
///
/// If you want to write to directory other than `.`, use `-w/--workdir` along with `-u` option.
///
/// ```
/// xlsx2csv input.xlsx -u -w test/
/// ```
///
/// The filename extension is detemined by delimiter, `,` to `.csv`, `\t` to `.tsv`, others will treat as ','.
///
/// By default, it will output all sheets, but if you want to select by sheet names with regex match, use `-I/--include` to include only matching, and `-X/--exclude` to exclude matching.
/// You could also combine these two option with *include-first-exclude-after* order.
///
/// ```
/// xlsx2csv input.xlsx -I '\S{3,}' -X 'Sheet'
/// ```
#[derive(Debug, Parser)]
pub struct Opt {
    /// Input Excel-like files, supports: .xls .xlsx .xlsb .xlsm .ods
    pub xlsx: PathBuf,
    /// Output each sheet to seprated file.
    ///
    /// If not setted, output first sheet to stdout.
    pub output: Vec<PathBuf>,
    /// List sheet names by id.
    #[clap(short, long, conflicts_with_all = &["output", "select", "use-sheet-names"])]
    pub list: bool,
    /// Use first line as header, which means use first line to select columns
    #[clap(short = 'U', long)]
    pub use_header: bool,
    /// Select sheet by name or id in output, only used when output to stdout.
    #[clap(short, long, conflicts_with = "output")]
    pub select: Option<SheetSelector>,
    /// Use sheet names as output filename prefix (in current dir or --workdir).
    #[clap(
        short,
        long,
        alias = "sheet",
        conflicts_with = "output",
        id = "use-sheet-names"
    )]
    pub use_sheet_names: bool,
    /// Output files location if `--use_sheet_names` setted
    #[clap(short, long, conflicts_with = "output", requires = "use-sheet-names")]
    pub workdir: Option<PathBuf>,
    /// A regex pattern for matching sheetnames to include, used with '-u'.
    #[clap(short = 'I', long, requires = "use-sheet-names")]
    pub include: Option<String>,
    /// A regex pattern for matching sheetnames to exclude, used with '-u'.
    #[clap(short = 'X', long, requires = "use-sheet-names")]
    pub exclude: Option<String>,
    /// Regex case insensitivedly.
    ///
    /// When this flag is provided, the include and exclude patterns will be searched case insensitively. used with '-u'.
    #[clap(short = 'i', long, requires = "use-sheet-names")]
    pub ignore_case: bool,
    /// Delimiter for output.
    ///
    /// If `use_sheet_names` setted, it will control the output filename extension: , -> csv, \t -> tsv
    #[clap(short, long, default_value = ",")]
    pub delimiter: Delimiter,
}

// list -> exit;
// use_sheet names -> sheet logic stuff;
// output.is_empty() -> (not output dir given) -> to_stdout;
//        // select -> to_stdout;
// fallback( sheet, output ) -> wtr as bytes -> run worksheet_to_csv(use header)

pub enum ActionType {
    List,
    UseSheetNames,
    OutputIsEmpty(InnerAction),
    DefaultToFile,
}

pub enum InnerAction {
    SelectValid(SheetSelector),
    SelectInvalid,
}

impl From<Opt> for InnerAction {
    fn from(opt: Opt) -> Self {
        if let Some(select) = opt.select.clone() {
            InnerAction::SelectValid(select)
        } else {
            InnerAction::SelectInvalid
        }
    }
}

impl From<&Opt> for InnerAction {
    fn from(opt: &Opt) -> Self {
        if let Some(select) = opt.select.clone() {
            InnerAction::SelectValid(select)
        } else {
            InnerAction::SelectInvalid
        }
    }
}

impl From<Opt> for ActionType {
    fn from(opt: Opt) -> Self {
        if opt.list {
            ActionType::List
        } else if opt.use_sheet_names {
            ActionType::UseSheetNames
        } else if opt.output.is_empty() {
            ActionType::OutputIsEmpty(opt.into())
        } else {
            ActionType::DefaultToFile
        }
    }
}

impl From<&Opt> for ActionType {
    fn from(opt: &Opt) -> Self {
        if opt.list {
            ActionType::List
        } else if opt.use_sheet_names {
            ActionType::UseSheetNames
        } else if opt.output.is_empty() {
            ActionType::OutputIsEmpty(Into::into(opt))
        } else {
            ActionType::DefaultToFile
        }
    }
}
