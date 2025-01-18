// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file.

use chrono::prelude::*;
use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Integer" => &[1, 2, 3, 4],
        "Float" => &[4.0, 5.0, 6.0, 7.0],
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            ],
        "Date" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
            ],
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
        ],
    )
    .unwrap();

    example1(&mut df).unwrap();
    example2(&df).unwrap();
}

// The ExcelWriter interface.
use polars_excel_writer::ExcelWriter;

fn example1(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file).finish(df)
}

// The PolarsXlsxWriter interface. For this simple case it is similar to the
// ExcelWriter interface but it has additional options to support more complex
// use cases.
use polars_excel_writer::PolarsXlsxWriter;

fn example2(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe2.xlsx")?;

    Ok(())
}
