// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframes to separate worksheets in an
//! Excel workbook.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    let df1: DataFrame = df!(
        "Data 1" => &[10, 11, 12, 13, 14, 15],
    )?;

    let df2: DataFrame = df!(
        "Data 2" => &[20, 21, 22, 23, 24, 25],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Write the first dataframe to the first/default worksheet.
    excel_writer.write_dataframe(&df1)?;

    // Add another worksheet and write the second dataframe to it.
    excel_writer.add_worksheet();
    excel_writer.write_dataframe(&df2)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
