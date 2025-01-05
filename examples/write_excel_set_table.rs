// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting properties of the worksheet table that wraps the output
//! dataframe.

use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::{Table, TableStyle};

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Add a `rust_xlsxwriter` table and set the style.
    let table = Table::new().set_style(TableStyle::Medium4);

    // Add the table to the Excel writer.
    xlsx_writer.set_table(&table);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
