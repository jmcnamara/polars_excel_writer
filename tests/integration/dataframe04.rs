// Test case that compares a file generated by polars_excel_writer with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

use crate::common;

use polars::prelude::*;
use polars_excel_writer::PolarsXlsxWriter;
use rust_xlsxwriter::XlsxError;

// Compare output against target Excel file using PolarsXlsxWriter.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let df: DataFrame = df!(
        "Foo" => &[1, 1, 1],
        "Bar" => &[2, 2, 2],
    )?;

    let mut xlsx_writer = PolarsXlsxWriter::new();

    xlsx_writer.write_dataframe_to_cell(&df, 0, 0)?;
    xlsx_writer.write_dataframe_to_cell(&df, 0, 3)?;

    xlsx_writer.save(filename)?;

    Ok(())
}

#[test]
fn dataframe_write_excel04() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe04")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}
