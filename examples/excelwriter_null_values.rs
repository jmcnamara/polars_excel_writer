// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting a value for Null values in the dataframe. The default
//! is to write them as blank cells.

use polars::prelude::*;

use polars_excel_writer::ExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a dataframe with Null values.
    let csv_string = "Foo,Bar\nNULL,B\nA,B\nA,NULL\nA,B\n";
    let buffer = std::io::Cursor::new(csv_string);
    let mut df = CsvReadOptions::default()
        .map_parse_options(|parse_options| {
            parse_options.with_null_values(Some(NullValues::AllColumnsSingle("NULL".into())))
        })
        .into_reader_with_file_handle(buffer)
        .finish()
        .unwrap();

    // Create a new file object.
    let mut file = std::fs::File::create("dataframe.xlsx").unwrap();

    // Write the dataframe to an Excel file using the Polars SerWriter
    // interface. This example also sets a string value for Null values.
    ExcelWriter::new(&mut file)
        .with_null_value("Null")
        .finish(&mut df)?;

    Ok(())
}
