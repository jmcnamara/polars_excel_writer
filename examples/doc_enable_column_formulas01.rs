// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates converting strings to formulas.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Region" => &["North", "South", "East", "West"],
        "Q1" => &[80, 20, 75, 85],
        "Q2" => &[80, 50, 65, 80],
        "Q3" => &[75, 60, 75, 80],
        "Q4" => &[70, 70, 65, 85],
        "Total" => &["=SUM(B2:E2)",
                     "=SUM(B3:E3)",
                     "=SUM(B4:E4)",
                     "=SUM(B5:E5)"],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Treat the Total column as a list of formulas.
    excel_writer.enable_column_formulas("Total");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
