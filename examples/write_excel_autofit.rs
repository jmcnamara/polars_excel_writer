// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates autofitting column widths in the output worksheet.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Col 1" => &["A", "B", "C", "D"],
        "Column 2" => &["A", "B", "C", "D"],
        "Column 3" => &["Hello", "World", "Hello, world", "Ciao"],
        "Column 4" => &[1234567, 12345678, 123456789, 1234567],
    )?;

    // Create a new Excel writer.
    let mut xlsx_writer = PolarsExcelWriter::new();

    // Autofit the output data.
    xlsx_writer.set_autofit(true);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
