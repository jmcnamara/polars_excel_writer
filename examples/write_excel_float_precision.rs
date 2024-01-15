// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! An example of writing a Polar Rust dataframe to an Excel file. This example
//! demonstrates how to set the precision of the float output. Setting the
//! precision to 3 is equivalent to an Excel number format of `0.000`.

use polars::prelude::*;

fn main() {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )
    .unwrap();

    example(&df).unwrap();
}

use polars_excel_writer::PolarsXlsxWriter;

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.set_float_precision(3);

    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
