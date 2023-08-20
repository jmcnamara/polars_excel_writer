// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing more than one Polar dataframes to an Excel worksheet.

use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    let df1: DataFrame = df!(
        "Data 1" => &[10, 20, 15, 25, 30, 20],
    )?;

    let df2: DataFrame = df!(
        "Data 2" => &[1.23, 2.34, 3.56],
    )?;

    // Write the dataframe to an Excel file.
    let mut writer = PolarsXlsxWriter::new();

    // Write two dataframes to the same worksheet.
    writer.write_dataframe_to_cell(&df1, 0, 0)?;
    writer.write_dataframe_to_cell(&df2, 0, 2)?;

    writer.write_excel("dataframe.xlsx")?;

    Ok(())
}
