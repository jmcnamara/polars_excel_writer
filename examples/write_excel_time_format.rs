// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to change the default format for Polars time types.

use chrono::prelude::*;
use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 00, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 18, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 37, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
        ],
    )
    .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.set_time_format("hh:mm");

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
