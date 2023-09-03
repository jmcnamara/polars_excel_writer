// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates autofitting column widths in the output worksheet.

use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Col 1" => &["A", "B", "C", "D"],
        "Column 2" => &["A", "B", "C", "D"],
        "Column 3" => &["Hello", "World", "Hello, world", "Ciao"],
        "Column 4" => &[1234567, 12345678, 123456789, 1234567],
    )
    .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.set_autofit(true);

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
