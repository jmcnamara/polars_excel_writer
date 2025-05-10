// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to set the precision of the float output. Setting the
//! precision to 3 is equivalent to an Excel number format of `0.000`.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsExcelWriter::new();

    // Set the float precision.
    xlsx_writer.set_float_precision(3);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
