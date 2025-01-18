// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars date types.

use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

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
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Set the date format.
    xlsx_writer.set_date_format("mmm d yyyy");

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
