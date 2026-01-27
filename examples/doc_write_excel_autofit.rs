// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates autofitting column widths in the output worksheet.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Col 1" => &["A", "B", "C", "D"],
        "Col 2" => &["Hello", "World", "Hello, world", "Ciao"],
        "Col 3" => &[1.234578, 123.45678, 123456.78, 12345679.0],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set an number format for column 3.
    excel_writer.set_column_format("Col 3", "$#,##0.00");

    // Autofit the output data.
    excel_writer.set_autofit(true);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
