// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting a value for Null values in the dataframe. The default
//! is to write them as blank cells.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
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

    // Write the dataframe to an Excel file.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set an output string value for Null.
    excel_writer.set_null_value("Null");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
