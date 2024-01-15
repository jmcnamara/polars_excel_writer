// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting properties of the worksheet table that wraps the output
//! dataframe.

use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )
    .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::*;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Add a `rust_xlsxwriter` table and set the style.
    let mut table = Table::new();
    table.set_style(TableStyle::Medium4);

    // Add the table to the Excel writer.
    xlsx_writer.set_table(&table);

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
