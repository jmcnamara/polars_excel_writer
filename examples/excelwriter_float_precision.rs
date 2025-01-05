// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to set the precision of the float output. Setting the
//! precision to 3 is equivalent to an Excel number format of `0.000`.

use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also adds a float precision.
    ExcelWriter::new(&mut file)
        .with_float_precision(3)
        .finish(&mut df)?;

    Ok(())
}
