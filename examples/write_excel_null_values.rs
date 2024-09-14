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
    let df = CsvReadOptions::default()
        .map_parse_options(|parse_options| {
            parse_options.with_null_values(Some(NullValues::AllColumnsSingle("NULL".into())))
        })
        .into_reader_with_file_handle(buffer)
        .finish()
        .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.set_null_value("Null");

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
