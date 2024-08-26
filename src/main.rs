// use calamine::DataType;
// use calamine::Reader;
// use calamine::{open_workbook_auto, Sheets};

mod cli;
mod delimiter;
mod sheet_selector;

use cli::{ActionType, InnerAction, Opt};
// use sheet_selector::SheetSelector;

use clap::Parser;
use regex::RegexBuilder;
use std::path::PathBuf;

fn worksheet_to_csv<W>(
    workbook: &ooxml::document::Workbook,
    sheet: &str,
    wtr: &mut csv::Writer<W>,
    header: bool,
) where
    W: std::io::Write,
{
    let worksheet = workbook
        .get_worksheet_by_name(sheet)
        .expect("worksheet name error");

    let mut iter = worksheet.rows();
    if header {
        let header = iter.next();
        if header.is_none() {
            return;
        }
        let header = header.unwrap();
        let size = header
            .into_iter()
            .position(|cell| cell.is_empty())
            .expect("find header row size");

        for row in worksheet.rows() {
            let cols: Vec<String> = row
                .into_iter()
                .take(size)
                .map(|cell| cell.to_string().unwrap_or_default())
                .collect();
            wtr.write_record(&cols).unwrap();
        }
    } else {
        for row in worksheet.rows() {
            let cols: Vec<String> = row
                .into_iter()
                .map(|cell| cell.to_string().unwrap_or_default())
                .collect();
            wtr.write_record(&cols).unwrap();
        }
    }
    wtr.flush().unwrap();
}

fn main() {
    let opt = Opt::parse();
    let xlsx = ooxml::document::SpreadsheetDocument::open(&opt.xlsx).expect("open xlsx file");
    let workbook = xlsx.get_workbook();
    //let mut workbook: Sheets = open_workbook_auto(&opt.xlsx).expect("open file");
    let sheetnames = workbook.worksheet_names();

    if sheetnames.is_empty() {
        panic!("input file has zero sheet!");
    }

    match ActionType::from(&opt) {
        ActionType::List => {
            if opt.list {
                for sheet in sheetnames {
                    println!("{}", sheet);
                }
                // return;
            }
        }

        ActionType::UseSheetNames => {
            if opt.use_sheet_names {
                let ignore_case = opt.ignore_case;
                let include_pattern = opt.include.map(|p| {
                    RegexBuilder::new(&p)
                        .case_insensitive(ignore_case)
                        .build()
                        .unwrap()
                });
                let exclude_pattern = opt.exclude.map(|p| {
                    RegexBuilder::new(&p)
                        .case_insensitive(ignore_case)
                        .build()
                        .unwrap()
                });
                let ext = opt.delimiter.to_file_extension();
                let workdir = opt.workdir.unwrap_or(PathBuf::new());
                for sheet in sheetnames
                    .iter()
                    .filter(|name| {
                        include_pattern
                            .as_ref()
                            .map(|r| r.is_match(name))
                            .unwrap_or(true)
                    })
                    .filter(|name| {
                        exclude_pattern
                            .as_ref()
                            .map(|r| !r.is_match(name))
                            .unwrap_or(true)
                    })
                {
                    let output = workdir.join(format!("{}.{}", sheet, ext));
                    println!("{}", output.display());
                    let mut wtr = csv::WriterBuilder::new()
                        .delimiter(opt.delimiter.as_byte())
                        .from_path(output)
                        .expect("open file for output");
                    worksheet_to_csv(workbook, sheet, &mut wtr, opt.use_header);
                }
            }
        }

        ActionType::OutputIsEmpty(inner) => {
            let stdout = std::io::stdout();
            let mut wtr = csv::WriterBuilder::new()
                .delimiter(opt.delimiter.as_byte())
                .from_writer(stdout);

            match inner {
                InnerAction::SelectValid(select) => {
                    let name = select.find_in(&sheetnames).expect("invalid selector");
                    worksheet_to_csv(workbook, name, &mut wtr, opt.use_header);
                }

                InnerAction::SelectInvalid => {
                    worksheet_to_csv(workbook, &sheetnames[0], &mut wtr, opt.use_header);
                }
            }
        }

        ActionType::DefaultToFile => {
            for (sheet, output) in sheetnames.iter().zip(opt.output.iter()) {
                println!("{}", output.display());
                let mut wtr = csv::WriterBuilder::new()
                    .delimiter(opt.delimiter.as_byte())
                    .from_path(output)
                    .expect("open file for output");
                worksheet_to_csv(workbook, sheet, &mut wtr, opt.use_header);
            }
        }
    }
}
