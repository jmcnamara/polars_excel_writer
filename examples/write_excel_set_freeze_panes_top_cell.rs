// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates freezing the top row and setting a non-default first row within
//! the pane.

use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Freeze the top row and set the first row in the range.
    xlsx_writer.set_freeze_panes(1, 0);
    xlsx_writer.set_freeze_panes_top_cell(3, 0);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
