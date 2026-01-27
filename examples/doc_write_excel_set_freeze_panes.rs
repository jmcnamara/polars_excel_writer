// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates freezing the top row.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Freeze the top row.
    excel_writer.set_freeze_panes(1, 0);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
