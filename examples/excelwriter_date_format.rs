// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars date types.

use chrono::prelude::*;
use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Date" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11),
            NaiveDate::from_ymd_opt(2023, 1, 12),
            NaiveDate::from_ymd_opt(2023, 1, 13),
            NaiveDate::from_ymd_opt(2023, 1, 14),
        ],
    )
    .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file)
        .with_date_format("mmm d yyyy")
        .finish(df)
}
