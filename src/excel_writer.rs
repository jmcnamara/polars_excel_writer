// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2025, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::collections::HashMap;
use std::io::{Seek, Write};
use std::path::Path;

use polars::prelude::*;
use polars_arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use rust_xlsxwriter::worksheet::IntoExcelData;
use rust_xlsxwriter::{Format, Table, TableColumn, Workbook, Worksheet};

/// `PolarsExcelWriter` provides an interface to serialize Polars dataframes to
/// Excel via the [`rust_xlsxwriter`] library. This allows Excel serialization
/// with a straightforward interface but also a high degree of configurability
/// over the output, when required.
///
/// For ease of use, and portability, `PolarsExcelWriter` tries to replicate the
/// interface options provided by the Polars Python [`write_excel()`] dataframe
/// method.
///
/// [`rust_xlsxwriter`]: ../../rust_xlsxwriter/
/// [`write_excel()`]:
///     https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel
///
/// ## Examples
///
/// Here is an example of writing a Polars Rust dataframe to an Excel file using
/// `PolarsExcelWriter`.
///
/// ```
/// # // This code is available in examples/doc_write_excel_intro.rs
/// #
/// use chrono::prelude::*;
/// use polars::prelude::*;
///
/// use polars_excel_writer::PolarsExcelWriter;
///
/// fn main() -> PolarsResult<()> {
///     // Create a sample dataframe for the example.
///     let df: DataFrame = df!(
///         "String" => &["North", "South", "East", "West"],
///         "Integer" => &[1, 2, 3, 4],
///         "Float" => &[4.0, 5.0, 6.0, 7.0],
///         "Time" => &[
///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
///             ],
///         "Date" => &[
///             NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
///             ],
///         "Datetime" => &[
///             NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
///             NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
///         ],
///     )?;
///
///     // Create a new Excel writer.
///     let mut excel_writer = PolarsExcelWriter::new();
///
///     // Write the dataframe to Excel.
///     excel_writer.write_dataframe(&df)?;
///
///     // Save the file to disk.
///     excel_writer.save("dataframe.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/excelwriter_intro.png">
///
/// ## Interacting with `rust_xlsxwriter`
///
/// The [`rust_xlsxwriter`] crate provides a fast native Rust interface for
/// creating Excel files with features such as formatting, formulas, charts,
/// hyperlinks, page setup, merged ranges, image support, rich multi-format
/// strings, autofilters and tables.
///
/// `PolarsExcelWriter` uses `rust_xlsxwriter` internally as its Excel writing
/// engine but it can also be used in conjunction with larger `rust_xlsxwriter`
/// programs to access functionality that it doesn't provide natively.
///
/// For example, say we wanted to write a dataframe to an Excel workbook but
/// also plot the data on an Excel chart. We can use `PolarsExcelWriter` crate
/// for the data writing part and `rust_xlsxwriter` for all the other
/// functionality.
///
/// Here is an example that demonstrate this:
///
/// [`rust_xlsxwriter`]: ../../rust_xlsxwriter/
///
/// ```
/// # // This code is available in examples/doc_write_excel_chart.rs
/// #
/// use polars::prelude::*;
/// use polars_excel_writer::PolarsExcelWriter;
/// use rust_xlsxwriter::{Chart, ChartType, Workbook};
///
/// fn main() -> PolarsResult<()> {
///     // Create a sample dataframe using `Polars`
///     let df: DataFrame = df!(
///         "Data" => &[10, 20, 15, 25, 30, 20],
///     )?;
///
///     // Get some dataframe dimensions that we will use for the chart range.
///     let row_min = 1; // Skip the header row.
///     let row_max = df.height() as u32;
///
///     // Create a new workbook and worksheet using `rust_xlsxwriter`.
///     let mut workbook = Workbook::new();
///     let worksheet = workbook.add_worksheet();
///
///     // Write the dataframe to the worksheet using `PolarsExcelWriter`.
///     let mut excel_writer = PolarsExcelWriter::new();
///     excel_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
///
///     // Move back to `rust_xlsxwriter` to create a new chart and have it plot the
///     // range of the dataframe in the worksheet.
///     let mut chart = Chart::new(ChartType::Line);
///     chart
///         .add_series()
///         .set_values(("Sheet1", row_min, 0, row_max, 0));
///
///     // Add the chart to the worksheet.
///     worksheet.insert_chart(0, 2, &chart)?;
///
///     // Save the file to disk.
///     workbook.save("chart.xlsx")?;
///
///     Ok(())
/// }
/// ```
///
/// Output file:
///
/// <img src="https://rustxlsxwriter.github.io/images/write_excel_chart.png">
///
/// The example demonstrates using three different crates to get the required
/// result:
///
/// 1. `polars` to create/manipulate the dataframe.
/// 2. `rust_xlsxwriter` to create an Excel workbook and worksheet and
///    optionally add other features to the worksheet.
/// 3. `polars_excel_writer::PolarsExcelWriter` to write the Polars dataframe to
///    the worksheet.
///
/// This may seem initially complicated but it divides the solution into
/// specialized libraries that are best suited for their task and it allow you
/// to access advanced Excel functionality that isn't provided by
/// `PolarsExcelWriter`.
///
/// One aspect to note in this example is the transparent handling of the
/// different error types. That is explained in the next section.
///
/// ## `PolarsError` and `XlsxError`
///
/// The `rust_xlsxwriter` crate uses an error type called
/// [`XlsxError`](rust_xlsxwriter::XlsxError) while `Polars` and
/// `PolarsExcelWriter` use an error type called [`PolarsError`]. In order to
/// make interoperability with Polars easier the `rust_xlsxwriter::XlsxError`
/// type maps to (and from) the `PolarsError` type.
///
/// That is why in the previous example we were able to use two different error
/// types within the same result/error context of `PolarsError`. The error type
/// is not explicit in the previous example but `PolarsResult<T>` expands to
/// `Result<T, PolarsError>`.
///
/// In order for this to be enabled you must use the `rust_xlsxwriter` `polars`
/// crate feature, however, this is turned on automatically when you use
/// `polars_excel_writer`.
///
pub struct PolarsExcelWriter {
    pub(crate) workbook: Workbook,
    pub(crate) options: WriterOptions,
}

impl Default for PolarsExcelWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl PolarsExcelWriter {
    /// Create a new `PolarsExcelWriter` instance.
    ///
    pub fn new() -> PolarsExcelWriter {
        let mut workbook = Workbook::new();
        workbook.add_worksheet();

        PolarsExcelWriter {
            workbook,
            options: WriterOptions::default(),
        }
    }

    /// Write a dataframe to a worksheet.
    ///
    /// Writes the supplied dataframe to cell `(0, 0)` in the first sheet of a
    /// new Excel workbook. See [`PolarsExcelWriter::write_dataframe_to_cell()`]
    /// below to write to a specific cell in the worksheet.
    ///
    /// The worksheet must be written to a file using
    /// [`save()`](PolarsExcelWriter::save).
    ///
    /// # Parameters
    ///
    /// - `df` - A Polars dataframe.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_write_dataframe.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Data" => &[10, 20, 15, 25, 30, 20],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_write_dataframe.png">
    ///
    pub fn write_dataframe(&mut self, df: &DataFrame) -> PolarsResult<()> {
        let options = self.options.clone();
        let worksheet = self.worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        Ok(())
    }

    /// Writes the supplied dataframe to a user defined cell in the first sheet
    /// of a new Excel workbook.
    ///
    /// Using this method it is possible to write more than one dataframe to the
    /// same worksheet, at different positions and without overlapping.
    ///
    /// The worksheet must be written to a file using
    /// [`save()`](PolarsExcelWriter::save).
    ///
    /// # Parameters
    ///
    /// - `df` - A Polars dataframe.
    /// - `row` - The zero indexed row number.
    /// - `col` - The zero indexed column number.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    /// # Examples
    ///
    /// An example of writing more than one Polar dataframes to an Excel
    /// worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_write_dataframe_to_cell.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df1: DataFrame = df!(
    ///         "Data 1" => &[10, 20, 15, 25, 30, 20],
    ///     )?;
    ///
    ///     let df2: DataFrame = df!(
    ///         "Data 2" => &[1.23, 2.34, 3.56],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Write two dataframes to the same worksheet.
    ///     excel_writer.write_dataframe_to_cell(&df1, 0, 0)?;
    ///     excel_writer.write_dataframe_to_cell(&df2, 0, 2)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_write_dataframe_to_cell.png">
    ///
    pub fn write_dataframe_to_cell(
        &mut self,
        df: &DataFrame,
        row: u32,
        col: u16,
    ) -> PolarsResult<()> {
        let options = self.options.clone();
        let worksheet = self.worksheet()?;

        Self::write_dataframe_internal(df, worksheet, row, col, &options)?;

        Ok(())
    }

    /// Write a dataframe to a user supplied worksheet.
    ///
    /// Writes the dataframe to a `rust_xlsxwriter` [`Worksheet`] object. This
    /// worksheet cannot be saved via
    /// [`save()`](PolarsExcelWriter::save). Instead it must be
    /// used in conjunction with a `rust_xlsxwriter` [`Workbook`].
    ///
    /// This is useful for mixing `PolarsExcelWriter` data writing with
    /// additional Excel functionality provided by `rust_xlsxwriter`. See
    /// [Interacting with `rust_xlsxwriter`](#interacting-with-rust_xlsxwriter)
    /// and the example below.
    ///
    /// # Parameters
    ///
    /// - `df` - A Polars dataframe.
    /// - `worksheet` - A `rust_xlsxwriter` [`Worksheet`].
    /// - `row` - The zero indexed row number.
    /// - `col` - The zero indexed column number.
    ///
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    /// # Examples
    ///
    /// An example of using `polars_excel_writer` in conjunction with
    /// `rust_xlsxwriter` to write a Polars dataframe to a worksheet and then
    /// add a chart to plot the data.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_chart.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsExcelWriter;
    /// use rust_xlsxwriter::{Chart, ChartType, Workbook};
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe using `Polars`
    ///     let df: DataFrame = df!(
    ///         "Data" => &[10, 20, 15, 25, 30, 20],
    ///     )?;
    ///
    ///     // Get some dataframe dimensions that we will use for the chart range.
    ///     let row_min = 1; // Skip the header row.
    ///     let row_max = df.height() as u32;
    ///
    ///     // Create a new workbook and worksheet using `rust_xlsxwriter`.
    ///     let mut workbook = Workbook::new();
    ///     let worksheet = workbook.add_worksheet();
    ///
    ///     // Write the dataframe to the worksheet using `PolarsExcelWriter`.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///     excel_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
    ///
    ///     // Move back to `rust_xlsxwriter` to create a new chart and have it plot the
    ///     // range of the dataframe in the worksheet.
    ///     let mut chart = Chart::new(ChartType::Line);
    ///     chart
    ///         .add_series()
    ///         .set_values(("Sheet1", row_min, 0, row_max, 0));
    ///
    ///     // Add the chart to the worksheet.
    ///     worksheet.insert_chart(0, 2, &chart)?;
    ///
    ///     // Save the file to disk.
    ///     workbook.save("chart.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_chart.png">
    ///
    pub fn write_dataframe_to_worksheet(
        &mut self,
        df: &DataFrame,
        worksheet: &mut Worksheet,
        row: u32,
        col: u16,
    ) -> PolarsResult<()> {
        let options = self.options.clone();

        Self::write_dataframe_internal(df, worksheet, row, col, &options)?;

        Ok(())
    }

    /// Save the Workbook as an xlsx file.
    ///
    /// The `save()` method writes all the workbook and worksheet data to
    /// a new xlsx file. It will overwrite any existing file.
    ///
    /// The method can be called multiple times so it is possible to get
    /// incremental files at different stages of a process, or to save the same
    /// Workbook object to different paths. However, `save()` is an
    /// expensive operation which assembles multiple files into an xlsx/zip
    /// container so for performance reasons you shouldn't call it
    /// unnecessarily.
    ///
    /// # Parameters
    ///
    /// - `path` - The path of the new Excel file to create as a `&str` or as a
    ///   [`std::path`] `Path` or `PathBuf` instance.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    pub fn save<P: AsRef<Path>>(&mut self, path: P) -> PolarsResult<()> {
        self.workbook.save(path)?;

        Ok(())
    }

    /// Turn on/off the dataframe header row in the Excel table. It is on by
    /// default.
    ///
    /// # Parameters
    ///
    /// - `has_header` - Export dataframe with/without header.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates saving the dataframe without a header.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_header.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "String" => &["North", "South", "East", "West"],
    ///         "Int" => &[1, 2, 3, 4],
    ///         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    ///     )
    ///     .unwrap();
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Turn off the default header.
    ///     excel_writer.set_header(false);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_has_header_off.png">
    ///
    pub fn set_header(&mut self, has_header: bool) -> &mut PolarsExcelWriter {
        let table = self.options.table.clone().set_header_row(has_header);
        self.options.table = table;
        self
    }

    /// Set an Excel format for a specific Polars data type.
    ///
    /// Sets a cell format to be applied to a Polar [`DataType`] type in a
    /// dataframe. The Polars' data types supported by Excel are:
    ///
    /// - [`DataType::Boolean`]
    /// - [`DataType::Int8`]
    /// - [`DataType::Int16`]
    /// - [`DataType::Int32`]
    /// - [`DataType::Int64`]
    /// - [`DataType::UInt8`]
    /// - [`DataType::UInt16`]
    /// - [`DataType::UInt32`]
    /// - [`DataType::UInt64`]
    /// - [`DataType::Float32`]
    /// - [`DataType::Float64`]
    /// - [`DataType::Date`]
    /// - [`DataType::Time`]
    /// - [`DataType::Datetime`]
    /// - [`DataType::String`]
    /// - [`DataType::Null`]
    ///
    /// **Formats**
    ///
    /// For more information on the formatting that is supported see the
    /// documentation for the `rust_xlsxwriter` [`Format`]. The majority of
    /// Excel cell formatting is available.
    ///
    /// See also the [Number Format Categories] section and the [Number Formats
    /// in different locales] sections in the `rust_xlsxwriter` documentation.
    ///
    /// [Number Format Categories]:
    ///     ../../rust_xlsxwriter/struct.Format.html#number-format-categories
    /// [Number Formats in different locales]:
    ///     ../../rust_xlsxwriter/struct.Format.html#number-formats-in-different-locales
    ///
    ///
    /// **Integer and Float types**
    ///
    /// Excel stores all integer and float types as [`f64`] floats without an
    /// explicit cell number format. It does, however, display them using a
    /// "printf"-like format of `%.16G` so that integers appear as integers and
    /// floats have the minimum required numbers of decimal places to maintain
    /// precision.
    ///
    /// Since there are many similar integer and float types in Polars, this
    /// library provides additional helper methods to set the format for related
    /// types:
    ///
    /// - [`set_dtype_int_format()`](PolarsExcelWriter::set_dtype_int_format):
    ///   All the integer types.
    /// - [`set_dtype_float_format()`](PolarsExcelWriter::set_dtype_float_format):
    ///   Both float types.
    /// - [`set_dtype_number_format()`](PolarsExcelWriter::set_dtype_number_format):
    ///   All the integer and float types.
    ///
    /// **Date and Time types**
    ///
    /// Datetimes in Excel are serial dates with days counted from an epoch
    /// (usually 1900-01-01) and the time is a percentage/decimal of the
    /// milliseconds in the day. Both the date and time are stored in the same
    /// `f64` value. For example, the date and time "2026/01/01 12:00:00" is
    /// stored as 46023.5.
    ///
    /// Datetimes in Excel must also be formatted with a number format like
    /// `"yyyy/mm/dd hh:mm"` or otherwise they will appear as numbers (which
    /// technically they are). By default the following formats are used for
    /// dates and times to match Polars `write_excel`:
    ///
    /// - Time: `hh:mm:ss;@`
    /// - Date: `yyyy-mm-dd;@`
    /// - Datetime: `yyyy-mm-dd hh:mm:ss`
    ///
    /// Alternative date and time formats can be specified as shown in the
    /// examples below and in the
    /// [`PolarsExcelWriter::set_dtype_datetime_format()`] method.
    ///
    /// Note, Excel doesn't use timezones or try to convert or encode timezone
    /// information in any way so they aren't supported by this library.
    ///
    /// # Parameters
    ///
    /// - `dtype` - A Polars [`DataType`] type.
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to change the default format for Polars time
    /// types.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_time_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Time" => &[
    ///             NaiveTime::from_hms_milli_opt(2, 00, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 18, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 37, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
    ///         ],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the time format.
    ///     excel_writer.set_dtype_format(DataType::Time, "hh:mm");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_time_format.png">
    ///
    /// This example demonstrates how to change the default format for Polars
    /// date types.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_date_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Date" => &[
    ///             NaiveDate::from_ymd_opt(2023, 1, 11),
    ///             NaiveDate::from_ymd_opt(2023, 1, 12),
    ///             NaiveDate::from_ymd_opt(2023, 1, 13),
    ///             NaiveDate::from_ymd_opt(2023, 1, 14),
    ///         ],
    ///     )?;
    ///
    ///     // Create a new Excel writer.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the date format.
    ///     excel_writer.set_dtype_format(DataType::Date, "mmm d yyyy");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_date_format.png">
    ///
    pub fn set_dtype_format(
        &mut self,
        dtype: DataType,
        format: impl Into<Format>,
    ) -> &mut PolarsExcelWriter {
        self.options.dtype_formats.insert(dtype, format.into());
        self
    }

    /// Set an Excel format for the Polars integer data types.
    ///
    /// Sets a cell format to be applied to Polar [`DataType`] integer types in
    /// a dataframe. This is a shortcut for setting the format for all the
    /// following integer types with
    /// [`set_dtype_format()`](PolarsExcelWriter::set_dtype_format):
    ///
    /// - [`DataType::Int8`]
    /// - [`DataType::Int16`]
    /// - [`DataType::Int32`]
    /// - [`DataType::Int64`]
    /// - [`DataType::UInt8`]
    /// - [`DataType::UInt16`]
    /// - [`DataType::UInt32`]
    /// - [`DataType::UInt64`]
    ///
    /// # Parameters
    ///
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    pub fn set_dtype_int_format(&mut self, format: impl Into<Format>) -> &mut PolarsExcelWriter {
        let format = format.into();

        self.set_dtype_format(DataType::Int8, format.clone());
        self.set_dtype_format(DataType::Int16, format.clone());
        self.set_dtype_format(DataType::Int32, format.clone());
        self.set_dtype_format(DataType::Int64, format.clone());
        self.set_dtype_format(DataType::UInt8, format.clone());
        self.set_dtype_format(DataType::UInt16, format.clone());
        self.set_dtype_format(DataType::UInt32, format.clone());
        self.set_dtype_format(DataType::UInt64, format.clone());

        self
    }

    /// Set an Excel format for the Polars float data types.
    ///
    /// Sets a cell format to be applied to Polar [`DataType`] float types in a
    /// dataframe. This method is a shortcut for setting the format for the
    /// following `f32`/`f64` float types with
    /// [`set_dtype_format()`](PolarsExcelWriter::set_dtype_format):
    ///
    /// - [`DataType::Float32`]
    /// - [`DataType::Float64`]
    ///
    /// The required format strings can be obtained from the `Format Cells ->
    /// Number` dialog in Excel.
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    /// # Parameters
    ///
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting an Excel number format for floats.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_float_format.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    ///     )
    ///     .unwrap();
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the float format.
    ///     excel_writer.set_dtype_float_format("#,##0.00");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_float_format.png">
    ///
    pub fn set_dtype_float_format(&mut self, format: impl Into<Format>) -> &mut PolarsExcelWriter {
        let format = format.into();

        self.set_dtype_format(DataType::Float32, format.clone());
        self.set_dtype_format(DataType::Float64, format.clone());

        self
    }

    /// Add a format for the Polars number data types.
    ///
    /// Sets a cell format to be applied to Polar [`DataType`] number types in a
    /// dataframe. This is a shortcut for setting the format for all the
    /// following types with
    /// [`set_dtype_format()`](PolarsExcelWriter::set_dtype_format):
    ///
    /// # Parameters
    ///
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// - [`DataType::Int8`]
    /// - [`DataType::Int16`]
    /// - [`DataType::Int32`]
    /// - [`DataType::Int64`]
    /// - [`DataType::UInt8`]
    /// - [`DataType::UInt16`]
    /// - [`DataType::UInt32`]
    /// - [`DataType::UInt64`]
    /// - [`DataType::Float32`]
    /// - [`DataType::Float64`]
    ///
    /// Note, excel treats all of these types as a [`f64`] float type.
    ///
    pub fn set_dtype_number_format(&mut self, format: impl Into<Format>) -> &mut PolarsExcelWriter {
        let format = format.into();

        self.set_dtype_int_format(format.clone());
        self.set_dtype_float_format(format.clone());
        self
    }

    /// Set an Excel format for the Polars datetime variants.
    ///
    /// Sets a cell format to be applied to Polar [`DataType::Datetime`]
    /// variants in a dataframe.
    ///
    /// The type signature for `DataType::Datetime` is `Datetime(TimeUnit,
    /// Option<TimeZone>)` with 3 possible [`TimeUnit`] variants and an optional
    /// [`TimeZone`] type.
    ///
    /// This method is a shortcut for setting the format for the following
    /// [`DataType::Datetime`] types with
    /// [`set_dtype_format()`](PolarsExcelWriter::set_dtype_format):
    ///
    /// - [`DataType::Datetime(TimeUnit::Nanoseconds, None)`]
    /// - [`DataType::Datetime(TimeUnit::Microseconds, None)`]
    /// - [`DataType::Datetime(TimeUnit::Milliseconds, None)`]
    ///
    /// Excel doesn't use timezones or try to convert or encode timezone
    /// information in any way so they aren't supported by this library.
    ///
    /// # Parameters
    ///
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to change the default format for Polars
    /// datetime types.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_datetime_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Datetime" => &[
    ///             NaiveDate::from_ymd_opt(2023, 1, 11).unwrap().and_hms_opt(1, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 12).unwrap().and_hms_opt(2, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 13).unwrap().and_hms_opt(3, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 14).unwrap().and_hms_opt(4, 0, 0).unwrap(),
    ///         ],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the datetime format.
    ///     excel_writer.set_dtype_datetime_format("hh::mm - mmm d yyyy");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_datetime_format.png">
    ///
    ///
    pub fn set_dtype_datetime_format(
        &mut self,
        format: impl Into<Format>,
    ) -> &mut PolarsExcelWriter {
        let format = format.into();

        self.set_dtype_format(
            DataType::Datetime(TimeUnit::Nanoseconds, None),
            format.clone(),
        );
        self.set_dtype_format(
            DataType::Datetime(TimeUnit::Microseconds, None),
            format.clone(),
        );
        self.set_dtype_format(
            DataType::Datetime(TimeUnit::Milliseconds, None),
            format.clone(),
        );

        self
    }

    /// Set the Excel number precision for floats.
    ///
    /// Set the number precision of all floats exported from the dataframe to
    /// Excel. The precision is converted to an Excel number format (see
    /// [`set_dtype_float_format()`](PolarsExcelWriter::set_dtype_float_format) above), so for
    /// example 3 is converted to the Excel format `0.000`.
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    /// # Parameters
    ///
    /// - `precision` - The floating point precision in the Excel range 1-30.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to set the precision of the float output. Setting the
    /// precision to 3 is equivalent to an Excel number format of `0.000`.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_float_precision.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    ///     )
    ///     .unwrap();
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the float precision.
    ///     excel_writer.set_float_precision(3);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_float_precision.png">
    ///
    pub fn set_float_precision(&mut self, precision: usize) -> &mut PolarsExcelWriter {
        if (1..=30).contains(&precision) {
            let precision = "0".repeat(precision);
            let format = Format::new().set_num_format(format!("0.{precision}"));
            self.set_dtype_float_format(format);
        }
        self
    }

    /// Add a format for a named column in the dataframe.
    ///
    /// Set an Excel format for a specific column in the dataframe. This is
    /// similar to the
    /// [`set_dtype_format()`](PolarsExcelWriter::set_dtype_format) method expect
    /// that is gives a different level of granularity. For example you could
    /// use this to format tow `f64` columns with different formats.
    ///
    /// # Parameters
    ///
    /// - `column_name` - The name of the column in the dataframe. Unknown
    ///   column names are silently ignored.
    /// - `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting formats for different columns.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_column_format.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "East" => &[1.0, 2.22, 3.333, 4.4444],
    ///         "West" => &[1.0, 2.22, 3.333, 4.4444],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the number formats for the columns.
    ///     excel_writer.set_column_format("East", "0.00");
    ///     excel_writer.set_column_format("West", "0.0000");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_column_format.png">
    ///
    ///
    pub fn set_column_format(
        &mut self,
        column_name: &str,
        format: impl Into<Format>,
    ) -> &mut PolarsExcelWriter {
        self.options
            .column_formats
            .insert(column_name.to_string(), format.into());
        self
    }

    /// Set the format for the header row.
    ///
    /// Set the format for the header row in the Excel table.
    ///
    /// # Parameters
    ///
    /// - `format` - A `rust_xlsxwriter` [`Format`].
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting the format for the header row.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_header_format.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    /// use rust_xlsxwriter::Format;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "East" => &[1, 1, 1, 1],
    ///         "West" => &[2, 2, 2, 2],
    ///         "North" => &[3, 3, 3, 3],
    ///         "South" => &[4, 4, 4, 4],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Create an set the header format.
    ///     let header_format = Format::new()
    ///         .set_background_color("#C6EFCE")
    ///         .set_font_color("#006100")
    ///         .set_bold();
    ///
    ///     // Set the number formats for the columns.
    ///     excel_writer.set_header_format(&header_format);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_header_format.png">
    ///
    pub fn set_header_format(&mut self, format: impl Into<Format>) -> &mut PolarsExcelWriter {
        self.options.header_format = Some(format.into());
        self
    }

    /// Replace Null values in the exported dataframe with string values.
    ///
    /// By default Null values in a dataframe aren't exported to Excel and will
    /// appear as empty cells. If you wish you can specify a string such as
    /// "Null", "NULL" or "N/A" as an alternative.
    ///
    /// # Parameters
    ///
    /// - `value` - A replacement string for Null values.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting a value for Null values in the dataframe. The
    /// default is to write them as blank cells.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_null_values.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a dataframe with Null values (represented as None).
    ///     let df = df! [
    ///         "Foo" => [None, Some("A"), Some("A"), Some("A")],
    ///         "Bar" => [Some("B"), Some("B"), None, Some("B")],
    ///     ]?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set an output string value for Null.
    ///     excel_writer.set_null_value("Null");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_null_values.png">
    ///
    pub fn set_null_value(&mut self, value: impl Into<String>) -> &mut PolarsExcelWriter {
        self.options.null_value = Some(value.into());
        self
    }

    /// Replace NaN values in the exported dataframe with string values.
    ///
    /// By default [`f64::NAN`] values in a dataframe are exported as the string
    /// "NAN" since Excel does not support NaN values.
    ///
    /// This method can be used to supply an alternative string value. See the
    /// example below.
    ///
    /// # Parameters
    ///
    /// - `value` - A replacement string for Null values.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates handling NaN and Infinity values with custom string
    /// representations.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_nan_value.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Default" => &["NAN", "INF", "-INF"],
    ///         "Custom" => &[f64::NAN, f64::INFINITY, f64::NEG_INFINITY],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set custom values for NaN, Infinity, and -Infinity.
    ///     excel_writer.set_nan_value("NaN");
    ///     excel_writer.set_infinity_value("Infinity");
    ///     excel_writer.set_neg_infinity_value("-Infinity");
    ///
    ///     // Autofit the output data, for clarity.
    ///     excel_writer.set_autofit(true);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_nan_value.png">
    ///
    pub fn set_nan_value(&mut self, value: impl Into<String>) -> &mut PolarsExcelWriter {
        self.options.nan_value = Some(value.into());
        self
    }

    /// Replace Infinity values in the exported dataframe with string values.
    ///
    /// By default [`f64::INFINITY`] values in a dataframe are exported as the
    /// string "INF" since Excel does not support Infinity values.
    ///
    /// This method can be used to supply an alternative string value. See the
    /// `set_nan_value()` example above.
    ///
    /// # Parameters
    ///
    /// - `value` - A replacement string for Null values.
    ///
    pub fn set_infinity_value(&mut self, value: impl Into<String>) -> &mut PolarsExcelWriter {
        self.options.infinity_value = Some(value.into());
        self
    }

    /// Replace Negative Infinity values in the exported dataframe with string
    /// values.
    ///
    /// By default [`f64::NEG_INFINITY`] values in a dataframe are exported as
    /// the string "-INF" since Excel does not support Infinity values.
    ///
    /// This method can be used to supply an alternative string value. See the
    /// `set_nan_value()` example above.
    ///
    /// # Parameters
    ///
    /// - `value` - A replacement string for Null values.
    ///
    pub fn set_neg_infinity_value(&mut self, value: impl Into<String>) -> &mut PolarsExcelWriter {
        self.options.neg_infinity_value = Some(value.into());
        self
    }

    /// Simulate autofit for columns in the dataframe output.
    ///
    /// Use a simulated autofit to adjust dataframe columns to the maximum
    /// string or number widths.
    ///
    /// **Note**: There are several limitations to this autofit method, see the
    /// `rust_xlsxwriter` docs on [`Worksheet::autofit()`] for details.
    ///
    /// [`Worksheet::autofit()`]:
    ///     ../../rust_xlsxwriter/worksheet/struct.Worksheet.html#method.autofit
    ///
    /// # Parameters
    ///
    /// - `autofit` - Turn autofit on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates autofitting column widths in the output worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_autofit.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Col 1" => &["A", "B", "C", "D"],
    ///         "Column 2" => &["A", "B", "C", "D"],
    ///         "Column 3" => &["Hello", "World", "Hello, world", "Ciao"],
    ///         "Column 4" => &[1234567, 12345678, 123456789, 1234567],
    ///     )?;
    ///
    ///     // Create a new Excel writer.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Autofit the output data.
    ///     excel_writer.set_autofit(true);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_autofit.png">
    ///
    pub fn set_autofit(&mut self, autofit: bool) -> &mut PolarsExcelWriter {
        self.options.use_autofit = autofit;
        self
    }

    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range `10 <= zoom <= 400`.
    ///
    /// # Parameters
    ///
    /// - `zoom` - The worksheet zoom level. The default zoom level is 100.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting the worksheet zoom level.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_zoom.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the worksheet zoom level.
    ///     excel_writer.set_zoom(200);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_zoom.png">
    ///
    pub fn set_zoom(&mut self, zoom: u16) -> &mut PolarsExcelWriter {
        self.options.zoom = zoom;
        self
    }

    /// Set the option to turn on/off the screen gridlines.
    ///
    /// The `set_screen_gridlines()` method is use to turn on/off gridlines on
    /// displayed worksheet. It is on by default.
    ///
    /// # Parameters
    ///
    /// - `enable` - Turn the property on/off. It is on by default.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates turning off the screen gridlines.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_screen_gridlines.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Turn off the screen gridlines.
    ///     excel_writer.set_screen_gridlines(false);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_screen_gridlines.png">
    ///
    pub fn set_screen_gridlines(&mut self, enable: bool) -> &mut PolarsExcelWriter {
        self.options.screen_gridlines = enable;

        self
    }

    /// Freeze panes in a worksheet.
    ///
    /// The `set_freeze_panes()` method can be used to divide a worksheet into
    /// horizontal or vertical regions known as panes and to “freeze” these
    /// panes so that the splitter bars are not visible.
    ///
    /// As with Excel the split is to the top and left of the cell. So to freeze
    /// the top row and leftmost column you would use `(1, 1)` (zero-indexed).
    ///
    /// You can set one of the row and col parameters as 0 if you do not want
    /// either the vertical or horizontal split. For example a common
    /// requirement is to freeze the top row which is done with the arguments
    /// `(1, 0)` see below.
    ///
    ///
    /// # Parameters
    ///
    /// - `row` - The zero indexed row number.
    /// - `col` - The zero indexed column number.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates freezing the top row.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_freeze_panes.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Freeze the top row.
    ///     excel_writer.set_freeze_panes(1, 0);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_freeze_panes.png">
    ///
    pub fn set_freeze_panes(&mut self, row: u32, col: u16) -> &mut PolarsExcelWriter {
        self.options.freeze_cell = (row, col);

        self
    }

    /// Set the top most cell in the scrolling area of a freeze pane.
    ///
    /// This method is used in conjunction with the
    /// [`PolarsExcelWriter::set_freeze_panes()`] method to set the top most
    /// visible cell in the scrolling range. For example you may want to freeze
    /// the top row but have the worksheet pre-scrolled so that a cell other
    /// than `(0, 0)` is visible in the scrolled area.
    ///
    /// # Parameters
    ///
    /// - `row` - The zero indexed row number.
    /// - `col` - The zero indexed column number.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates freezing the top row and setting a non-default first row
    /// within the pane.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_freeze_panes_top_cell.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Freeze the top row and set the first row in the range.
    ///     excel_writer.set_freeze_panes(1, 0);
    ///     excel_writer.set_freeze_panes_top_cell(3, 0);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_freeze_panes_top_cell.png">
    ///
    pub fn set_freeze_panes_top_cell(&mut self, row: u32, col: u16) -> &mut PolarsExcelWriter {
        self.options.top_cell = (row, col);

        self
    }

    /// Turn on/off the autofilter for the table header.
    ///
    /// By default Excel adds an autofilter to the header of a table. This
    /// method can be used to turn it off if necessary.
    ///
    /// Note, you can call this method directly on a [`Table`] object which is
    /// passed to [`PolarsExcelWriter::set_table()`].
    ///
    /// # Parameters
    ///
    /// - `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn set_autofilter(&mut self, enable: bool) -> &mut PolarsExcelWriter {
        let table = self.options.table.clone().set_autofilter(enable);
        self.options.table = table;

        self
    }

    /// Set the worksheet table for the output dataframe.
    ///
    /// By default, and by convention with the Polars [`write_excel()`] method,
    /// `PolarsExcelWriter` adds an Excel worksheet table to each exported
    /// dataframe.
    ///
    /// Tables in Excel are a way of grouping a range of cells into a single
    /// entity that has common formatting or that can be referenced from
    /// formulas. Tables can have column headers, autofilters, total rows,
    /// column formulas and different formatting styles.
    ///
    /// The image below shows a default table in Excel with the default
    /// properties shown in the ribbon bar.
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/table_intro.png">
    ///
    /// The `set_table()` method allows you to pass a pre-configured
    /// `rust_xlsxwriter` table and override any of the default [`Table`]
    /// properties.
    ///
    /// [`write_excel()`]:
    ///     https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel
    ///
    ///
    /// # Parameters
    ///
    /// - `table` - A `rust_xlsxwriter` [`Table`] reference.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting properties of the worksheet table that wraps the
    /// output dataframe.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_table.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// # use rust_xlsxwriter::{Table, TableStyle};
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Add a `rust_xlsxwriter` table and set the style.
    ///     let table = Table::new().set_style(TableStyle::Medium4);
    ///
    ///     // Add the table to the Excel writer.
    ///     excel_writer.set_table(&table);
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_table.png">
    ///
    pub fn set_table(&mut self, table: &Table) -> &mut PolarsExcelWriter {
        self.options.table = table.clone();
        self
    }

    /// Set the worksheet name for the output dataframe.
    ///
    /// Set the name of the worksheet that the dataframe is written to. If the
    /// name isn't set then it will be the default Excel name of `Sheet1` (or
    /// `Sheet2`, `Sheet3`, etc. if more than one worksheet is added).
    ///
    /// # Parameters
    ///
    /// - `name` - The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    ///   - The name must be less than 32 characters.
    ///   - The name cannot be blank.
    ///   - The name cannot contain any of the characters: `[ ] : * ? / \`.
    ///   - The name cannot start or end with an apostrophe.
    ///   - The name shouldn't be "History" (case-insensitive) since that is
    ///     reserved by Excel.
    ///   - It must not be a duplicate of another worksheet name used in the
    ///     workbook.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    /// Excel has several rules that govern what a worksheet name can be. See
    /// [`set_name()` errors] for more details.
    ///
    /// [`set_name()` errors]:
    ///     ../../rust_xlsxwriter/worksheet/struct.Worksheet.html#errors
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting the name for the output worksheet.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_set_worksheet_name.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Set the worksheet name.
    ///     excel_writer.set_worksheet_name("Polars Data")?;
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_worksheet_name.png">
    ///
    pub fn set_worksheet_name(
        &mut self,
        name: impl Into<String>,
    ) -> PolarsResult<&mut PolarsExcelWriter> {
        let worksheet = self.worksheet()?;
        worksheet.set_name(name)?;
        Ok(self)
    }

    /// Add a new worksheet to the output workbook.
    ///
    /// Add a worksheet to the workbook so that dataframes can be written to
    /// more than one worksheet. This is useful when you have several dataframes
    /// that you wish to have on separate worksheets.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframes to separate worksheets in
    /// an Excel workbook.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_add_worksheet.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsExcelWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     let df1: DataFrame = df!(
    ///         "Data 1" => &[10, 11, 12, 13, 14, 15],
    ///     )?;
    ///
    ///     let df2: DataFrame = df!(
    ///         "Data 2" => &[20, 21, 22, 23, 24, 25],
    ///     )?;
    ///
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Write the first dataframe to the first/default worksheet.
    ///     excel_writer.write_dataframe(&df1)?;
    ///
    ///     // Add another worksheet and write the second dataframe to it.
    ///     excel_writer.add_worksheet();
    ///     excel_writer.write_dataframe(&df2)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_add_worksheet.png">
    ///
    pub fn add_worksheet(&mut self) -> &mut PolarsExcelWriter {
        self.workbook.add_worksheet();

        self
    }

    /// Get the current worksheet in the workbook.
    ///
    /// Get a reference to the current/last worksheet in the workbook in order
    /// to manipulate it with a `rust_xlsxwriter` [`Worksheet`] method. This is
    /// occasionally useful when you need to access some feature of the
    /// worksheet APIs that isn't supported directly by `PolarsExcelWriter`.
    ///
    /// Note, it is also possible to create a [`Worksheet`] separately and then
    /// write the Polar dataframe to it using the
    /// [`write_dataframe_to_worksheet()`](PolarsExcelWriter::write_dataframe_to_worksheet)
    /// method. That latter is more useful if you need to do a lot of
    /// manipulation of the worksheet.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates getting a reference to the worksheet used to write the
    /// dataframe and setting its tab color.
    ///
    /// ```
    /// # // This code is available in examples/doc_write_excel_worksheet.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # use polars_excel_writer::PolarsExcelWriter;
    /// #
    /// # fn main() -> PolarsResult<()> {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )?;
    /// #
    ///     // Write the dataframe to an Excel file.
    ///     let mut excel_writer = PolarsExcelWriter::new();
    ///
    ///     // Get the worksheet that the dataframe will be written to.
    ///     let worksheet = excel_writer.worksheet()?;
    ///
    ///     // Set the tab color for the worksheet using a `rust_xlsxwriter` worksheet
    ///     // method.
    ///     worksheet.set_tab_color("#FF9900");
    ///
    ///     // Write the dataframe to Excel.
    ///     excel_writer.write_dataframe(&df)?;
    ///
    ///     // Save the file to disk.
    ///     excel_writer.save("dataframe.xlsx")?;
    /// #
    /// #     Ok(())
    /// # }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_worksheet.png">
    ///
    pub fn worksheet(&mut self) -> PolarsResult<&mut Worksheet> {
        let mut last_index = self.workbook.worksheets().len();

        // Add a worksheet if there isn't one already.
        if last_index == 0 {
            self.workbook.add_worksheet();
        } else {
            last_index -= 1;
        }

        let worksheet = self.workbook.worksheet_from_index(last_index)?;

        Ok(worksheet)
    }

    /// Method to support writing to `ExcelWriter` writer<W>.
    ///
    /// This is a hidden method to support the deprecated `ExcelWriter` module.
    /// Support for `ExcelWriter` may be moved to a separate crate in the
    /// future.
    ///
    /// # Parameters
    ///
    /// - `df` - A Polars dataframe.
    /// - `writer` - A generic type that supports the Write trait.
    ///
    /// # Errors
    ///
    /// A [`PolarsError::ComputeError`] that wraps a `rust_xlsxwriter`
    /// [`XlsxError`](rust_xlsxwriter::XlsxError) error.
    ///
    #[doc(hidden)]
    pub fn save_to_writer<W>(&mut self, df: &DataFrame, writer: W) -> PolarsResult<()>
    where
        W: Write + Seek + Send,
    {
        let options = self.options.clone();
        let worksheet = self.worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        self.workbook.save_to_writer(writer)?;

        Ok(())
    }

    // -----------------------------------------------------------------------
    // Internal functions/methods.
    // -----------------------------------------------------------------------

    // Write the dataframe to a `rust_xlsxwriter` Worksheet. It is structured as
    // an associated method to allow it to handle external worksheets.
    #[allow(clippy::too_many_lines)]
    fn write_dataframe_internal(
        df: &DataFrame,
        worksheet: &mut Worksheet,
        row_offset: u32,
        col_offset: u16,
        options: &WriterOptions,
    ) -> Result<(), PolarsError> {
        let header_offset = u32::from(options.table.has_header_row());
        let mut table_columns = vec![];

        // Set NaN and Infinity values, if required.
        if let Some(nan_value) = &options.nan_value {
            worksheet.set_nan_value(nan_value);
        }
        if let Some(infinity_value) = &options.infinity_value {
            worksheet.set_infinity_value(infinity_value);
        }
        if let Some(neg_infinity_value) = &options.neg_infinity_value {
            worksheet.set_neg_infinity_value(neg_infinity_value);
        }

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col = col_offset + col_num as u16;

            // Add the header format to the table columns
            if let Some(header_format) = &options.header_format {
                let table_column = TableColumn::new().set_header_format(header_format);

                table_columns.push(table_column);
            }

            // Store the column names for use as table headers.
            if options.table.has_header_row() {
                worksheet.write(row_offset, col, column.name().as_str())?;
            }

            // Check for a custom dtype or column format.
            let mut format = None;
            if let Some(dtype_format) = options.dtype_formats.get(column.dtype()) {
                format = Some(dtype_format);
            }

            // Column format takes precedence over dtype format since it is more specific.
            if let Some(column_format) = options.column_formats.get(&column.name().to_string()) {
                format = Some(column_format);
            }

            // Write the row data for each column/type.
            for (row_num, any_value) in column.as_materialized_series().iter().enumerate() {
                let row = header_offset + row_offset + row_num as u32;

                // Map Polars AnyValue types to Excel/rust_xlsxwriter types.
                match any_value {
                    // Write the number types to the worksheet.
                    AnyValue::Int8(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::Int16(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::Int32(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::Int64(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::UInt8(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::UInt16(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::UInt32(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::UInt64(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::Float32(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::Float64(value) => write_value(worksheet, row, col, value, format)?,

                    // Write the string types to the worksheet.
                    AnyValue::String(value) => write_value(worksheet, row, col, value, format)?,
                    AnyValue::StringOwned(value) => {
                        write_value(worksheet, row, col, value.as_str(), format)?;
                    }

                    AnyValue::Datetime(value, time_units, _) => {
                        let value = match time_units {
                            TimeUnit::Nanoseconds => timestamp_ns_to_datetime(value),
                            TimeUnit::Microseconds => timestamp_us_to_datetime(value),
                            TimeUnit::Milliseconds => timestamp_ms_to_datetime(value),
                        };

                        write_value(worksheet, row, col, &value, format)?;
                        worksheet.set_column_width(col, 18)?;
                    }

                    AnyValue::Date(value) => {
                        let value = date32_to_date(value);

                        write_value(worksheet, row, col, &value, format)?;
                        worksheet.set_column_width(col, 10)?;
                    }

                    AnyValue::Time(value) => {
                        let value = time64ns_to_time(value);

                        write_value(worksheet, row, col, &value, format)?;
                    }

                    // Write the boolean type to the worksheet.
                    AnyValue::Boolean(value) => write_value(worksheet, row, col, value, format)?,

                    // Write null type to the worksheet.
                    AnyValue::Null => {
                        if let Some(value) = &options.null_value {
                            // Use user defined null value.
                            write_value(worksheet, row, col, value, format)?;
                        } else if format.is_some() {
                            // If a format is set then write a blank cell.
                            write_value(worksheet, row, col, "", format)?;
                        }
                    }

                    _ => {
                        polars_bail!(
                            ComputeError:
                            "Polars AnyValue data type '{}' is not supported by Excel",
                            any_value.dtype()
                        );
                    }
                }
            }
        }

        // Create a table for the dataframe range.
        let (mut max_row, max_col) = df.shape();
        if !options.table.has_header_row() {
            max_row -= 1;
        }
        if options.table.has_total_row() {
            max_row += 1;
        }

        // Add a column header format via table columns.
        let mut table = options.table.clone();
        if !table_columns.is_empty() {
            table = table.set_columns(&table_columns);
        }

        // Add the table to the worksheet.
        worksheet.add_table(
            row_offset,
            col_offset,
            row_offset + max_row as u32,
            col_offset + max_col as u16 - 1,
            &table,
        )?;

        // Autofit the columns.
        if options.use_autofit {
            worksheet.autofit();
        }

        // Set the zoom level.
        worksheet.set_zoom(options.zoom);

        // Set the screen gridlines.
        worksheet.set_screen_gridlines(options.screen_gridlines);

        // Set the worksheet panes.
        worksheet.set_freeze_panes(options.freeze_cell.0, options.freeze_cell.1)?;
        worksheet.set_freeze_panes_top_cell(options.top_cell.0, options.top_cell.1)?;

        Ok(())
    }
}

// Generic function to write a Polars typed value to the worksheet with an
// optional format.
fn write_value(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: impl IntoExcelData,
    format: Option<&Format>,
) -> Result<(), PolarsError> {
    match format {
        Some(format) => worksheet.write_with_format(row, col, value, format)?,
        None => worksheet.write(row, col, value)?,
    };

    Ok(())
}

// -----------------------------------------------------------------------
// Helper structs.
// -----------------------------------------------------------------------

/// Backwards compatibility type alias for the deprecated `PolarsXlsxWriter`
/// struct name.
#[deprecated(since = "0.15.0", note = "use `PolarsExcelWriter` instead")]
pub type PolarsXlsxWriter = PolarsExcelWriter;

// A struct for storing and passing configuration settings.
#[derive(Clone)]
pub(crate) struct WriterOptions {
    pub(crate) use_autofit: bool,
    pub(crate) null_value: Option<String>,
    pub(crate) nan_value: Option<String>,
    pub(crate) infinity_value: Option<String>,
    pub(crate) neg_infinity_value: Option<String>,
    pub(crate) table: Table,
    pub(crate) zoom: u16,
    pub(crate) screen_gridlines: bool,
    pub(crate) freeze_cell: (u32, u16),
    pub(crate) top_cell: (u32, u16),
    pub(crate) header_format: Option<Format>,
    pub(crate) column_formats: HashMap<String, Format>,
    pub(crate) dtype_formats: HashMap<DataType, Format>,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl WriterOptions {
    fn new() -> WriterOptions {
        WriterOptions {
            use_autofit: false,
            null_value: None,
            nan_value: None,
            infinity_value: None,
            neg_infinity_value: None,
            table: Table::new(),
            zoom: 100,
            screen_gridlines: true,
            freeze_cell: (0, 0),
            top_cell: (0, 0),
            header_format: None,
            column_formats: HashMap::new(),
            dtype_formats: HashMap::from([
                (DataType::Time, "hh:mm:ss;@".into()),
                (DataType::Date, "yyyy\\-mm\\-dd;@".into()),
                (
                    DataType::Datetime(TimeUnit::Nanoseconds, None),
                    "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
                ),
                (
                    DataType::Datetime(TimeUnit::Microseconds, None),
                    "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
                ),
                (
                    DataType::Datetime(TimeUnit::Milliseconds, None),
                    "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
                ),
            ]),
        }
    }
}
