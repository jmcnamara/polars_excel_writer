// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates saving the dataframe with a header (which is the default).

use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also turns off the default header.
    ExcelWriter::new(&mut file)
        .has_header(false)
        .finish(&mut df)?;

    Ok(())
}
