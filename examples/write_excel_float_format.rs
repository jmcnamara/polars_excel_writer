// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting an Excel number format for floats.

use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Set the float format.
    xlsx_writer.set_float_format("#,##0.00");

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
