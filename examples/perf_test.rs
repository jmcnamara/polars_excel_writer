// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

//! Simple performance test to compare with the Python Polars example in
//! `perf_test.py`.

use chrono::prelude::*;
use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;
use std::time::Instant;

const DATA_SIZE: usize = 250_000;

fn main() {
    // Create a sample dataframe for testing.
    let df: DataFrame = df!(
        "Int" => &[1; DATA_SIZE],
        "Float" => &[123.456789; DATA_SIZE],
        "Date" => &[Utc::now().date_naive(); DATA_SIZE],
        "String" => &["Test"; DATA_SIZE],
    )
    .unwrap();

    let timer = Instant::now();
    example(&df).unwrap();
    println!("Elapsed time: {:.2?}", timer.elapsed());
}

fn example(df: &DataFrame) -> PolarsResult<()> {
    let mut xlsx_writer = PolarsXlsxWriter::new();
    xlsx_writer.write_dataframe(df)?;
    xlsx_writer.save("dataframe_rs.xlsx")?;

    Ok(())
}
