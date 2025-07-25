// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting a value for Null values in the dataframe. The
//! default is to write them as blank cells.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a dataframe with Null values (represented as None).
    let df = df! [
        "Foo" => [None, Some("A"), Some("A"), Some("A")],
        "Bar" => [Some("B"), Some("B"), None, Some("B")],
    ]?;

    // Write the dataframe to an Excel file.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set an output string value for Null.
    excel_writer.set_null_value("Null");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
