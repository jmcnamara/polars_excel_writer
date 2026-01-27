// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file and returning
//! it as a byte vector buffer.

use std::fs::File;
use std::io::Write;

use polars::prelude::*;
use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Month" => &["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
        "Volume" => &[100, 110, 100, 90, 90, 105],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the dataframe as an excel file in a byte vector buffer.
    let buf = excel_writer.save_to_buffer()?;

    // Write the buffer to a file for the sake of the example.
    let mut file = File::create("dataframe.xlsx")?;
    Write::write_all(&mut file, &buf)?;

    Ok(())
}
