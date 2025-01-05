// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file.

use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Data" => &[10, 20, 15, 25, 30, 20],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
