// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This
//! demonstrates setting the format for the header row.

use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::Format;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "East" => &[1, 1, 1, 1],
        "West" => &[2, 2, 2, 2],
        "North" => &[3, 3, 3, 3],
        "South" => &[4, 4, 4, 4],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Create an set the header format.
    let header_format = Format::new()
        .set_background_color("#C6EFCE")
        .set_font_color("#006100")
        .set_bold();

    xlsx_writer.set_header_format(&header_format);

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
