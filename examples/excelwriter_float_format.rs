// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting an Excel number format for floats.

use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    )?;

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also adds a float format.
    ExcelWriter::new(&mut file)
        .with_float_format("#,##0.00")
        .finish(&mut df)?;

    Ok(())
}
