// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates autofitting column widths in the output worksheet.

use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "Col 1" => &["A", "B", "C", "D"],
        "Column 2" => &["A", "B", "C", "D"],
        "Column 3" => &["Hello", "World", "Hello, world", "Ciao"],
        "Column 4" => &[1234567, 12345678, 123456789, 1234567],
    )
    .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file).with_autofit().finish(df)
}
