// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars date types.

use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Date" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11),
            NaiveDate::from_ymd_opt(2023, 1, 12),
            NaiveDate::from_ymd_opt(2023, 1, 13),
            NaiveDate::from_ymd_opt(2023, 1, 14),
        ],
    )?;

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also adds a date format.
    ExcelWriter::new(&mut file)
        .with_date_format("mmm d yyyy")
        .finish(&mut df)?;

    Ok(())
}
