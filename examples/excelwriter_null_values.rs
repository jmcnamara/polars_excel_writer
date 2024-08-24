// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting a value for Null values in the dataframe. The default
//! is to write them as blank cells.

use polars::prelude::*;

fn main() {
    // Create a dataframe with Null values.
    let csv_string = "Foo,Bar\nNULL,B\nA,B\nA,NULL\nA,B\n";
    let buffer = std::io::Cursor::new(csv_string);
    let mut df = CsvReadOptions::default()
        .map_parse_options(|parse_options| {
            parse_options.with_null_values(Some(NullValues::AllColumnsSingle("NULL".to_string())))
        })
        .into_reader_with_file_handle(buffer)
        .finish()
        .unwrap();

    example(&mut df).unwrap();
}

use polars_excel_writer::ExcelWriter;

fn example(df: &mut DataFrame) -> PolarsResult<()> {
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    ExcelWriter::new(&mut file)
        .with_null_value("Null")
        .finish(df)
}
