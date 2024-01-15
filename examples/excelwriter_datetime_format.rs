// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars datetime types.

use chrono::prelude::*;
use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 12).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 13).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 14).unwrap().and_hms_opt(4, 0, 0).unwrap(),
        ],
    )
    .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file)
        .with_datetime_format("hh::mm - mmm d yyyy")
        .finish(df)
}
