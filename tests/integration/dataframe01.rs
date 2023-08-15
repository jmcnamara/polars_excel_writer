// Test case that compares a file generated by polars_excel_writer with a file
// created by Excel.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2023, John McNamara, jmcnamara@cpan.org

use crate::common;

use polars::prelude::*;
use polars_excel_writer::ExcelWriter;
use rust_xlsxwriter::XlsxError;

// Test case to compare dataframe output against an Excel file.
fn create_new_xlsx_file(filename: &str) -> Result<(), XlsxError> {
    let mut df: DataFrame = df!(
        "Foo" => &[1, 1, 1],
        "Bar" => &[2, 2, 2],
    )
    .unwrap();

    let mut file = std::fs::File::create(filename).unwrap();

    ExcelWriter::new(&mut file).finish(&mut df).unwrap();

    Ok(())
}

#[test]
fn dataframe_excelwriter01() {
    let test_runner = common::TestRunner::new()
        .set_name("dataframe01")
        .set_function(create_new_xlsx_file)
        .initialize();

    test_runner.assert_eq();
    test_runner.cleanup();
}