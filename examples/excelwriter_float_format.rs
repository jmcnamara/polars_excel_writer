// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting an Excel number format for floats.

use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let mut df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    )
    .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file)
        .with_float_format("#,##0.00")
        .finish(df)
}
