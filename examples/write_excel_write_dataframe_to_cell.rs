// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! An example of writing more than one Polar dataframes to an Excel worksheet.

use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df1: DataFrame = df!(
        "Data 1" => &[10, 20, 15, 25, 30, 20],
    )?;

    let df2: DataFrame = df!(
        "Data 2" => &[1.23, 2.34, 3.56],
    )?;

    // Write the dataframe to an Excel file.
    let mut xlsx_writer = PolarsExcelWriter::new();

    // Write two dataframes to the same worksheet.
    xlsx_writer.write_dataframe_to_cell(&df1, 0, 0)?;
    xlsx_writer.write_dataframe_to_cell(&df2, 0, 2)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
