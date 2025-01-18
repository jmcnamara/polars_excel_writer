// Entry point for `polars_excel_writer` library.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

//! A crate for serializing Polars dataframes to Excel Xlsx files.
//!
//! The `polars_excel_writer` crate provides a primary interface
//! [`PolarsXlsxWriter`] which is a configurable Excel serializer that resembles
//! the interface options provided by the Polars [`write_excel()`] dataframe
//! method.
//!
//! It also provides a secondary [`ExcelWriter`] interface which is a simpler
//! Excel serializer that implements the Polars [`SerWriter`] trait to write a
//! dataframe to an Excel Xlsx file. However, unless you have existing code that
//! uses the [`SerWriter`] trait you should use the [`PolarsXlsxWriter`]
//! interface.
//!
//! This crate uses [`rust_xlsxwriter`] to do the Excel serialization.
//!
//! [`SerWriter`]:
//!     https://docs.rs/polars/latest/polars/prelude/trait.SerWriter.html
//!
//![`write_excel()`]:
//!    https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel
//!
//!  # Examples
//!
//! An example of writing a Polar Rust dataframe to an Excel file:.
//!
//! ```rust
//! # // This code is available in examples/write_excel_intro.rs
//! #
//! # use chrono::prelude::*;
//! # use polars::prelude::*;
//! #
//! use polars_excel_writer::PolarsXlsxWriter;
//!
//! fn main() -> PolarsResult<()> {
//!     // Create a sample dataframe for the example.
//!     let df: DataFrame = df!(
//!         "String" => &["North", "South", "East", "West"],
//!         "Integer" => &[1, 2, 3, 4],
//!         "Float" => &[4.0, 5.0, 6.0, 7.0],
//!         "Time" => &[
//!             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
//!             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
//!             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
//!             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
//!             ],
//!         "Date" => &[
//!             NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
//!             ],
//!         "Datetime" => &[
//!             NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
//!             NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
//!         ],
//!     )
//!     .unwrap();
//!
//!     // Create a new Excel writer.
//!     let mut xlsx_writer = PolarsXlsxWriter::new();
//!
//!     // Write the dataframe to Excel.
//!     xlsx_writer.write_dataframe(&df)?;
//!
//!     // Save the file to disk.
//!     xlsx_writer.save("dataframe.xlsx")?;
//!
//!     Ok(())
//! }
//! ```
//!
//! Output file:
//!
//! <img src="https://rustxlsxwriter.github.io/images/write_excel_combined.png">
//!
//!
//! ## Performance
//!
//! The table below shows the performance of writing a dataframe using Python
//! Polars, Python Pandas and `PolarsXlsxWriter`.
//!
//! - Performance data:
//!
//!   | Test Case                     | Time (s) | Relative (%) |
//!   | :---------------------------- | :------- | :----------- |
//!   | `Polars`                      |     6.49 |         100% |
//!   | `Pandas`                      |    10.92 |         168% |
//!   | `polars_excel_writer`         |     1.22 |          19% |
//!   | `polars_excel_writer` + `zlib`|     1.08 |          17% |
//!
//! The tested configurations were:
//!
//! - `Polars`: The dataframe was created in Python Polars and written using the
//!   [`write_excel()`] function. See [`perf_test.py`].
//! - `Pandas`: The dataframe was created in Polars but converted to Pandas and
//!   then written via the Pandas [`to_excel()`] function. See also
//!   [`perf_test.py`].
//! - `polars_excel_writer`: The dataframe was created in Rust Polars and
//!   written using the `PolarsXlsxWriter` interface. See [`perf_test.rs`].
//! - `polars_excel_writer` + `zlib`: Same as the previous test case but uses
//!   the `zlib` feature flag to enable the C zlib library in conjunction with
//!   the backend `ZipWriter`.
//!
//! **Note**: The performance was tested for the dataframe writing code only.
//! The code used to create the dataframes was omitted from the test results.
//!
//! [`perf_test.py`]:
//! https://github.com/jmcnamara/polars_excel_writer/blob/main/examples/perf_test.py
//! [`perf_test.rs`]:
//! https://github.com/jmcnamara/polars_excel_writer/blob/main/examples/perf_test.rs
//! [`to_excel()`]:
//! https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.to_excel.html

/// A module that exports the `PolarsXlsxWriter` struct which provides the
/// primary Excel Xlsx serializer that works with Polars dataframes and which
/// can also interact with the [`rust_xlsxwriter`] writing engine that it wraps.
pub mod xlsx_writer;

/// A module that exports the `ExcelWriter` struct which implements the
/// secondary Polars `SerWriter` trait to serialize a dataframe to an Excel Xlsx
/// file. Use [`xlsx_writer`] unless you have existing code that uses the
/// `SerWriter` trait.
pub mod write;

#[doc(hidden)]
pub use write::*;
#[doc(hidden)]
pub use xlsx_writer::*;

pub use ExcelWriter;
pub use PolarsXlsxWriter;
