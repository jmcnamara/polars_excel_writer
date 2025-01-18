// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars time types.

use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 00, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 18, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 37, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
        ],
    )?;

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also adds a time format.
    ExcelWriter::new(&mut file)
        .with_time_format("hh:mm")
        .finish(&mut df)?;

    Ok(())
}
