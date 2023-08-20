// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file.

use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    let df: DataFrame = df!(
        "Data" => &[10, 20, 15, 25, 30, 20],
    )?;

    // Write the dataframe to an Excel file.
    let mut writer = PolarsXlsxWriter::new();

    writer.write_dataframe(&df)?;
    writer.write_excel("dataframe.xlsx")?;

    Ok(())
}
