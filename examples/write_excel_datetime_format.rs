// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars datetime types.

use chrono::prelude::*;
use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 12).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 13).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2023, 1, 14).unwrap().and_hms_opt(4, 0, 0).unwrap(),
        ],
    )
    .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut writer = PolarsXlsxWriter::new();

    writer.set_datetime_format("hh::mm - mmm d yyyy");

    writer.write_dataframe(df)?;
    writer.write_excel("dataframe.xlsx")?;

    Ok(())
}
