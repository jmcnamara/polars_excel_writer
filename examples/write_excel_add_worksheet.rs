// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframes to separate worksheets in an
//! Excel workbook.

use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    let df1: DataFrame = df!(
        "Data 1" => &[10, 11, 12, 13, 14, 15],
    )?;

    let df2: DataFrame = df!(
        "Data 2" => &[20, 21, 22, 23, 24, 25],
    )?;

    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Write the first dataframe to the first/default worksheet.
    xlsx_writer.write_dataframe(&df1)?;

    // Add another worksheet and write the second dataframe to it.
    xlsx_writer.add_worksheet();
    xlsx_writer.write_dataframe(&df2)?;

    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
