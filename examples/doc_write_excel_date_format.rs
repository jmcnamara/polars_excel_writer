// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars date types.

use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Date" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11),
            NaiveDate::from_ymd_opt(2023, 1, 12),
            NaiveDate::from_ymd_opt(2023, 1, 13),
            NaiveDate::from_ymd_opt(2023, 1, 14),
        ],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the date format.
    excel_writer.set_dtype_format(DataType::Date, "mmm d yyyy");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
