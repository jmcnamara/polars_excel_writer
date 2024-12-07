// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023-2024, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::io::{Seek, Write};
use std::path::Path;

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, Workbook, Worksheet};

/// `PolarsXlsxWriter` provides an Excel Xlsx serializer that works with Polars
/// dataframes and which can also interact with the [`rust_xlsxwriter`] writing
/// engine that it wraps. This allows simple Excel serialization of worksheets
/// with a straightforward interface but also a high degree of configurability
/// over the output when required.
///
/// It is a complimentary interface to the much simpler
/// [`ExcelWriter`](crate::ExcelWriter) which implements the Polars
/// [`SerWriter`] trait to serialize dataframes, and which is also part of this
/// crate.
///
/// `ExcelWriter` and `PolarsXlsxWriter` both use the [`rust_xlsxwriter`] crate.
/// The `rust_xlsxwriter` library can only create new files. It cannot read or
/// modify existing files.
///
/// [`rust_xlsxwriter`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/
///
/// `PolarsXlsxWriter` tries to replicate the interface options provided by the
///  Polars Python [`write_excel()`] dataframe method.
///
/// [`write_excel()`]:
///     https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel
///
/// ## Examples
///
/// Here is an example of writing a Polars Rust dataframe to an Excel file using
/// `PolarsXlsxWriter`.
///
/// ```
/// # // This code is available in examples/excelwriter_intro.rs
/// #
/// use chrono::prelude::*;
/// use polars::prelude::*;
///
/// fn main() {
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
///     )
///     .unwrap();
///
///     example(&df).unwrap();
/// }
///
/// use polars_excel_writer::PolarsXlsxWriter;
///
/// fn example(df: &DataFrame) -> PolarsResult<()> {
///     let mut xlsx_writer = PolarsXlsxWriter::new();
///
///     xlsx_writer.write_dataframe(df)?;
///     xlsx_writer.save("dataframe.xlsx")?;
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
/// `PolarsXlsxWriter` uses `rust_xlsxwriter` internally as its Excel writing
/// engine but it can also be used in conjunction with larger `rust_xlsxwriter`
/// programs to access functionality that it doesn't provide natively.
///
/// For example, say we wanted to write a dataframe to an Excel workbook but
/// also plot the data on an Excel chart. We can use `PolarsXlsxWriter` crate
/// for the data writing part and `rust_xlsxwriter` for all the other
/// functionality.
///
/// Here is an example that demonstrate this:
///
/// [`rust_xlsxwriter`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/
///
/// ```
/// # // This code is available in examples/write_excel_chart.rs
/// #
/// use polars::prelude::*;
/// use polars_excel_writer::PolarsXlsxWriter;
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
///     // Write the dataframe to the worksheet using `PolarsXlsxWriter`.
///     let mut xlsx_writer = PolarsXlsxWriter::new();
///     xlsx_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
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
/// - 1: `polars` to create/manipulate the dataframe.
/// - 2a: `rust_xlsxwriter` to create an Excel workbook and worksheet.
/// - 3: `polars_excel_writer::PolarsXlsxWriter` to write the Polars dataframe
///   to the worksheet.
/// - 2b: `rust_xlsxwriter` to add other features to the worksheet.
///
/// This may seem initially complicated but it divides the solution into
/// specialized libraries that are best suited for their task and it allow you
/// to access advanced Excel functionality that isn't provided by
/// `PolarsXlsxWriter`.
///
/// One aspect to note in this example is the transparent handling of the
/// different error types. That is explained in the next section.
///
/// ## `PolarsError` and `XlsxError`
///
/// The `rust_xlsxwriter` crate uses an error type called
/// [`XlsxError`](rust_xlsxwriter::XlsxError) while `Polars` and
/// `PolarsXlsxWriter` use an error type called [`PolarsError`]. In order to
/// make interoperability with Polars easier the `rust_xlsxwriter::XlsxError`
/// type maps to (and from) the `PolarsError` type.
///
/// That is why in the previous example we were able to use two different error
/// types within the same result/error context of `PolarsError`. Note, the error
/// type is not explicit in the previous example but `PolarsResult<T>` expands
/// to `Result<T, PolarsError>`.
///
/// In order for this to be enabled you must use the `rust_xlsxwriter` `polars`
/// crate feature, however, this is turned on automatically when you use
/// `polars_excel_writer`.
///
pub struct PolarsXlsxWriter {
    pub(crate) workbook: Workbook,
    pub(crate) options: WriterOptions,
}

impl Default for PolarsXlsxWriter {
    fn default() -> Self {
        Self::new()
    }
}

impl PolarsXlsxWriter {
    /// Create a new `PolarsXlsxWriter` instance.
    ///
    pub fn new() -> PolarsXlsxWriter {
        let mut workbook = Workbook::new();
        workbook.add_worksheet();

        PolarsXlsxWriter {
            workbook,
            options: WriterOptions::default(),
        }
    }

    /// Write a dataframe to a worksheet.
    ///
    /// Writes the supplied dataframe to cell `(0, 0)` in the first sheet of a
    /// new Excel workbook.
    ///
    /// The worksheet must be written to a file using
    /// [`save()`](PolarsXlsxWriter::save).
    ///
    /// # Parameters
    ///
    /// * `df` - A Polars dataframe.
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
    /// # // This code is available in examples/write_excel_write_dataframe.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     let df: DataFrame = df!(
    ///         "Data" => &[10, 20, 15, 25, 30, 20],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.write_dataframe(&df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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

    /// Write a dataframe to a user defined cell in a worksheet.
    ///
    /// Writes the supplied dataframe to a user defined cell in the first sheet
    /// of a new Excel workbook.
    ///
    /// Since the dataframe can be positioned within the worksheet it is possible
    /// to write more than one to the same worksheet (without overlapping).
    ///
    /// The worksheet must be written to a file using
    /// [`save()`](PolarsXlsxWriter::save).
    ///
    /// # Parameters
    ///
    /// * `df` - A Polars dataframe.
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
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
    /// # // This code is available in examples/write_excel_write_dataframe_to_cell.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn main() -> PolarsResult<()> {
    ///     let df1: DataFrame = df!(
    ///         "Data 1" => &[10, 20, 15, 25, 30, 20],
    ///     )?;
    ///
    ///     let df2: DataFrame = df!(
    ///         "Data 2" => &[1.23, 2.34, 3.56],
    ///     )?;
    ///
    ///     // Write the dataframe to an Excel file.
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     // Write two dataframes to the same worksheet.
    ///     xlsx_writer.write_dataframe_to_cell(&df1, 0, 0)?;
    ///     xlsx_writer.write_dataframe_to_cell(&df2, 0, 2)?;
    ///
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    /// [`save()`](PolarsXlsxWriter::save). Instead it must be
    /// used in conjunction with a `rust_xlsxwriter` [`Workbook`].
    ///
    /// This is useful for mixing `PolarsXlsxWriter` data writing with
    /// additional Excel functionality provided by `rust_xlsxwriter`. See
    /// [Interacting with `rust_xlsxwriter`](#interacting-with-rust_xlsxwriter)
    /// and the example below.
    ///
    /// # Parameters
    ///
    /// * `df` - A Polars dataframe.
    /// * `worksheet` - A `rust_xlsxwriter` [`Worksheet`].
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
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
    /// # // This code is available in examples/write_excel_chart.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsXlsxWriter;
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
    ///     // Write the dataframe to the worksheet using `PolarsXlsxWriter`.
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///     xlsx_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
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
    /// * `path` - The path of the new Excel file to create as a `&str` or as a
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

    /// Turn on/off the dataframe header in the exported Excel file.
    ///
    /// Turn on/off the dataframe header row in the Excel table. It is on by
    /// default.
    ///
    /// # Parameters
    ///
    /// * `has_header` - Export dataframe with/without header.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates saving the dataframe without a header.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_header.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_header(false);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_header(&mut self, has_header: bool) -> &mut PolarsXlsxWriter {
        let table = self.options.table.clone().set_header_row(has_header);
        self.options.table = table;
        self
    }

    /// Set the Excel number format for time values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default time format used by this library is
    /// `hh:mm:ss;@`. This method can be used to specify an alternative user
    /// defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// # Parameters
    ///
    /// * `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to change the default format for Polars time types.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_time_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Time" => &[
    ///             NaiveTime::from_hms_milli_opt(2, 00, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 18, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 37, 3, 456).unwrap(),
    ///             NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
    ///         ],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_time_format("hh:mm");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_time_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.time_format = format.into();
        self
    }

    /// Set the Excel number format for date values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default date format used by this library is
    /// `yyyy-mm-dd;@`. This method can be used to specify an alternative user
    /// defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// # Parameters
    ///
    /// * `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to change the default format for Polars date types.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_date_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Date" => &[
    ///             NaiveDate::from_ymd_opt(2023, 1, 11),
    ///             NaiveDate::from_ymd_opt(2023, 1, 12),
    ///             NaiveDate::from_ymd_opt(2023, 1, 13),
    ///             NaiveDate::from_ymd_opt(2023, 1, 14),
    ///         ],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_date_format("mmm d yyyy");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/excelwriter_date_format.png">
    ///
    pub fn set_date_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.date_format = format.into();
        self
    }

    /// Set the Excel number format for datetime values.
    ///
    /// [Datetimes in Excel] are stored as f64 floats with a format used to
    /// display them. The default datetime format used by this library is
    /// `yyyy-mm-dd hh:mm:ss`. This method can be used to specify an alternative
    /// user defined format.
    ///
    /// [Datetimes in Excel]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.ExcelDateTime.html#datetimes-in-excel
    ///
    /// # Parameters
    ///
    /// * `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to change the default format for Polars datetime types.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_datetime_format.rs
    /// #
    /// use chrono::prelude::*;
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Datetime" => &[
    ///             NaiveDate::from_ymd_opt(2023, 1, 11).unwrap().and_hms_opt(1, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 12).unwrap().and_hms_opt(2, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 13).unwrap().and_hms_opt(3, 0, 0).unwrap(),
    ///             NaiveDate::from_ymd_opt(2023, 1, 14).unwrap().and_hms_opt(4, 0, 0).unwrap(),
    ///         ],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_datetime_format("hh::mm - mmm d yyyy");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_datetime_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.datetime_format = format.into();
        self
    }

    /// Set the Excel number format for floats.
    ///
    /// Set the Excel number format for f32/f64 float types using an Excel
    /// number format string. These format strings can be obtained from the
    /// `Format Cells -> Number` dialog in Excel.
    ///
    /// See the [Number Format Categories] section and subsequent Number Format
    /// sections in the `rust_xlsxwriter` documentation.
    ///
    /// [Number Format Categories]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#number-format-categories
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    /// # Parameters
    ///
    /// * `format` - A `rust_xlsxwriter` [`Format`] or an Excel number format
    ///   string that can be converted to a `Format`.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting an Excel number format for floats.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_float_format.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_float_format("#,##0.00");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_float_format(&mut self, format: impl Into<Format>) -> &mut PolarsXlsxWriter {
        self.options.float_format = format.into();
        self
    }

    /// Set the Excel number precision for floats.
    ///
    /// Set the number precision of all floats exported from the dataframe to
    /// Excel. The precision is converted to an Excel number format (see
    /// [`set_float_format()`](PolarsXlsxWriter::set_float_format) above), so for
    /// example 3 is converted to the Excel format `0.000`.
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    /// # Parameters
    ///
    /// * `precision` - The floating point precision in the Excel range 1-30.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to set the precision of the float output. Setting the
    /// precision to 3 is equivalent to an Excel number format of `0.000`.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_float_precision.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_float_precision(3);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_float_precision(&mut self, precision: usize) -> &mut PolarsXlsxWriter {
        if (1..=30).contains(&precision) {
            let precision = "0".repeat(precision);
            self.options.float_format = Format::new().set_num_format(format!("0.{precision}"));
        }
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
    /// * `null_value` - A replacement string for Null values.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting a value for Null values in the dataframe. The default
    /// is to write them as blank cells.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_null_values.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a dataframe with Null values.
    /// #     let csv_string = "Foo,Bar\nNULL,B\nA,B\nA,NULL\nA,B\n";
    /// #     let buffer = std::io::Cursor::new(csv_string);
    /// #     let df = CsvReadOptions::default()
    /// #         .map_parse_options(|parse_options| {
    /// #             parse_options.with_null_values(Some(NullValues::AllColumnsSingle("NULL".into())))
    /// #         })
    /// #         .into_reader_with_file_handle(buffer)
    /// #         .finish()
    /// #         .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_null_value("Null");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_null_value(&mut self, null_value: impl Into<String>) -> &mut PolarsXlsxWriter {
        self.options.null_string = Some(null_value.into());
        self
    }

    /// Simulate autofit for columns in the dataframe output.
    ///
    /// Use a simulated autofit to adjust dataframe columns to the maximum
    /// string or number widths.
    ///
    /// **Note**: There are several limitations to this autofit method, see the
    /// `rust_xlsxwriter` docs on [`worksheet.autofit()`] for details.
    ///
    /// [`worksheet.autofit()`]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#method.autofit
    ///
    /// # Parameters
    ///
    /// * `autofit` - Turn autofit on/off. It is off by default.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates autofitting column widths in the output worksheet.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_autofit.rs
    /// #
    /// use polars::prelude::*;
    ///
    /// fn main() {
    ///     // Create a sample dataframe for the example.
    ///     let df: DataFrame = df!(
    ///         "Col 1" => &["A", "B", "C", "D"],
    ///         "Column 2" => &["A", "B", "C", "D"],
    ///         "Column 3" => &["Hello", "World", "Hello, world", "Ciao"],
    ///         "Column 4" => &[1234567, 12345678, 123456789, 1234567],
    ///     )
    ///     .unwrap();
    ///
    ///     example(&df).unwrap();
    /// }
    ///
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_autofit(true);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn set_autofit(&mut self, autofit: bool) -> &mut PolarsXlsxWriter {
        self.options.use_autofit = autofit;
        self
    }

    /// Set the worksheet zoom factor.
    ///
    /// Set the worksheet zoom factor in the range 10 <= zoom <= 400.
    ///
    /// # Parameters
    ///
    /// * `zoom` - The worksheet zoom level. The default zoom level is 100.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting the worksheet zoom level.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_zoom.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_zoom(200);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_zoom.png">
    ///
    pub fn set_zoom(&mut self, zoom: u16) -> &mut PolarsXlsxWriter {
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
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates turning off the screen gridlines.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_screen_gridlines.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_screen_gridlines(false);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img src="https://rustxlsxwriter.github.io/images/write_excel_set_screen_gridlines.png">
    ///
    pub fn set_screen_gridlines(&mut self, enable: bool) -> &mut PolarsXlsxWriter {
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
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates freezing the top row.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_freeze_panes.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_freeze_panes(1, 0);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_freeze_panes.png">
    ///
    pub fn set_freeze_panes(&mut self, row: u32, col: u16) -> &mut PolarsXlsxWriter {
        self.options.freeze_cell = (row, col);

        self
    }

    /// Set the top most cell in the scrolling area of a freeze pane.
    ///
    /// This method is used in conjunction with the
    /// [`PolarsXlsxWriter::set_freeze_panes()`] method to set the top most
    /// visible cell in the scrolling range. For example you may want to freeze
    /// the top row but have the worksheet pre-scrolled so that a cell other
    /// than `(0, 0)` is visible in the scrolled area.
    ///
    /// # Parameters
    ///
    /// * `row` - The zero indexed row number.
    /// * `col` - The zero indexed column number.
    ///
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates freezing the top row and setting a non-default first row
    /// within the pane.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_freeze_panes_top_cell.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_freeze_panes(1, 0);
    ///     xlsx_writer.set_freeze_panes_top_cell(3, 0);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_freeze_panes_top_cell.png">
    ///
    pub fn set_freeze_panes_top_cell(&mut self, row: u32, col: u16) -> &mut PolarsXlsxWriter {
        self.options.top_cell = (row, col);

        self
    }

    /// Turn on/off the autofilter for the table header.
    ///
    /// By default Excel adds an autofilter to the header of a table. This
    /// method can be used to turn it off if necessary.
    ///
    /// Note, you can call this method directly on a [`Table`] object which is
    /// passed to [`PolarsXlsxWriter::set_table()`].
    ///
    /// # Parameters
    ///
    /// * `enable` - Turn the property on/off. It is on by default.
    ///
    pub fn set_autofilter(&mut self, enable: bool) -> &mut PolarsXlsxWriter {
        let table = self.options.table.clone().set_autofilter(enable);
        self.options.table = table;

        self
    }

    /// Set the worksheet table for the output dataframe.
    ///
    /// By default, and by convention with the Polars [`write_excel()`] method,
    /// `PolarsXlsxWriter` adds an Excel worksheet table to each exported
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
    /// * `table` - A `rust_xlsxwriter` [`Table`] reference.
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting properties of the worksheet table that wraps the
    /// output dataframe.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_table.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    /// use rust_xlsxwriter::*;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     // Add a `rust_xlsxwriter` table and set the style.
    ///     let table = Table::new().set_style(TableStyle::Medium4);
    ///
    ///     // Add the table to the Excel writer.
    ///     xlsx_writer.set_table(&table);
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
    /// ```
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/write_excel_set_table.png">
    ///
    pub fn set_table(&mut self, table: &Table) -> &mut PolarsXlsxWriter {
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
    /// * `name` - The worksheet name. It must follow the Excel rules, shown
    ///   below.
    ///
    ///   * The name must be less than 32 characters.
    ///   * The name cannot be blank.
    ///   * The name cannot contain any of the characters: `[ ] : * ? / \`.
    ///   * The name cannot start or end with an apostrophe.
    ///   * The name shouldn't be "History" (case-insensitive) since that is
    ///     reserved by Excel.
    ///   * It must not be a duplicate of another worksheet name used in the
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
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/worksheet/struct.Worksheet.html#errors
    ///
    /// # Examples
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting the name for the output worksheet.
    ///
    /// ```
    /// # // This code is available in examples/write_excel_set_worksheet_name.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     xlsx_writer.set_worksheet_name("Polars Data")?;
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
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
    ) -> PolarsResult<&mut PolarsXlsxWriter> {
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
    /// # // This code is available in examples/write_excel_add_worksheet.rs
    /// #
    /// # use polars::prelude::*;
    /// use polars_excel_writer::PolarsXlsxWriter;
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
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     // Write the first dataframe to the first/default worksheet.
    ///     xlsx_writer.write_dataframe(&df1)?;
    ///
    ///     // Add another worksheet and write the second dataframe to it.
    ///     xlsx_writer.add_worksheet();
    ///     xlsx_writer.write_dataframe(&df2)?;
    ///
    ///     xlsx_writer.save("dataframe.xlsx")?;
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
    pub fn add_worksheet(&mut self) -> &mut PolarsXlsxWriter {
        self.workbook.add_worksheet();

        self
    }

    /// Get the current worksheet in the workbook.
    ///
    /// Get a reference to the current/last worksheet in the workbook in order
    /// to manipulate it with a `rust_xlsxwriter` [`Worksheet`] method. This is
    /// occasionally useful when you need to access some feature of the
    /// worksheet APIs that isn't supported directly by `PolarsXlsxWriter`.
    ///
    /// Note, it is also possible to create a [`Worksheet`] separately and then
    /// write the Polar dataframe to it using the
    /// [`write_dataframe_to_worksheet()`](PolarsXlsxWriter::write_dataframe_to_worksheet)
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
    /// # // This code is available in examples/write_excel_worksheet.rs
    /// #
    /// # use polars::prelude::*;
    /// #
    /// # fn main() {
    /// #     // Create a sample dataframe for the example.
    /// #     let df: DataFrame = df!(
    /// #         "String" => &["North", "South", "East", "West"],
    /// #         "Int" => &[1, 2, 3, 4],
    /// #         "Float" => &[1.0, 2.22, 3.333, 4.4444],
    /// #     )
    /// #     .unwrap();
    /// #
    /// #     example(&df).unwrap();
    /// # }
    /// #
    /// use polars_excel_writer::PolarsXlsxWriter;
    ///
    /// fn example(df: &DataFrame) -> PolarsResult<()> {
    ///     let mut xlsx_writer = PolarsXlsxWriter::new();
    ///
    ///     // Get the worksheet that the dataframe will be written to.
    ///     let worksheet = xlsx_writer.worksheet()?;
    ///
    ///     // Set the tab color for the worksheet using a `rust_xlsxwriter` worksheet
    ///     // method.
    ///     worksheet.set_tab_color("#FF9900");
    ///
    ///     xlsx_writer.write_dataframe(df)?;
    ///     xlsx_writer.save("dataframe.xlsx")?;
    ///
    ///     Ok(())
    /// }
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

    // -----------------------------------------------------------------------
    // Internal functions/methods.
    // -----------------------------------------------------------------------

    // Method to support writing to ExcelWriter writer<W>.
    pub(crate) fn save_to_writer<W>(&mut self, df: &DataFrame, writer: W) -> PolarsResult<()>
    where
        W: Write + Seek + Send,
    {
        let options = self.options.clone();
        let worksheet = self.worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        self.workbook.save_to_writer(writer)?;

        Ok(())
    }

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

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_offset + col_num as u16;

            // Store the column names for use as table headers.
            if options.table.has_header_row() {
                worksheet.write(row_offset, col_num, column.name().as_str())?;
            }

            // Write the row data for each column/type.
            for (row_num, data) in column.as_materialized_series().iter().enumerate() {
                let row_num = header_offset + row_offset + row_num as u32;

                // Map the Polars Series AnyValue types to Excel/rust_xlsxwriter
                // types.
                match data {
                    AnyValue::Int8(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt8(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Int16(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt16(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Int32(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::UInt32(value) => {
                        worksheet.write_number(row_num, col_num, value)?;
                    }
                    AnyValue::Int64(value) => {
                        // Allow u64 conversion within Excel's limits.
                        #[allow(clippy::cast_precision_loss)]
                        worksheet.write_number(row_num, col_num, value as f64)?;
                    }
                    AnyValue::UInt64(value) => {
                        // Allow u64 conversion within Excel's limits.
                        #[allow(clippy::cast_precision_loss)]
                        worksheet.write_number(row_num, col_num, value as f64)?;
                    }
                    AnyValue::Float32(value) => {
                        worksheet.write_number_with_format(
                            row_num,
                            col_num,
                            value,
                            &options.float_format,
                        )?;
                    }
                    AnyValue::Float64(value) => {
                        worksheet.write_number_with_format(
                            row_num,
                            col_num,
                            value,
                            &options.float_format,
                        )?;
                    }
                    AnyValue::String(value) => {
                        worksheet.write_string(row_num, col_num, value)?;
                    }
                    AnyValue::StringOwned(value) => {
                        worksheet.write_string(row_num, col_num, value.as_str())?;
                    }
                    AnyValue::Boolean(value) => {
                        worksheet.write_boolean(row_num, col_num, value)?;
                    }
                    AnyValue::Null => {
                        if let Some(null_string) = &options.null_string {
                            worksheet.write_string(row_num, col_num, null_string)?;
                        }
                    }
                    AnyValue::Datetime(value, time_units, _) => {
                        let datetime = match time_units {
                            TimeUnit::Nanoseconds => timestamp_ns_to_datetime(value),
                            TimeUnit::Microseconds => timestamp_us_to_datetime(value),
                            TimeUnit::Milliseconds => timestamp_ms_to_datetime(value),
                        };
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            datetime,
                            &options.datetime_format,
                        )?;
                        worksheet.set_column_width(col_num, 18)?;
                    }
                    AnyValue::Date(value) => {
                        let date = date32_to_date(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            date,
                            &options.date_format,
                        )?;
                        worksheet.set_column_width(col_num, 10)?;
                    }
                    AnyValue::Time(value) => {
                        let time = time64ns_to_time(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            time,
                            &options.time_format,
                        )?;
                    }
                    _ => {
                        polars_bail!(
                            ComputeError:
                            "Polars AnyValue data type '{}' is not supported by Excel",
                            data.dtype()
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

        // Add the table to the worksheet.
        worksheet.add_table(
            row_offset,
            col_offset,
            row_offset + max_row as u32,
            col_offset + max_col as u16 - 1,
            &options.table,
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

// -----------------------------------------------------------------------
// Helper structs.
// -----------------------------------------------------------------------

// A struct for storing and passing configuration settings.
#[derive(Clone)]
pub(crate) struct WriterOptions {
    pub(crate) use_autofit: bool,
    pub(crate) date_format: Format,
    pub(crate) time_format: Format,
    pub(crate) float_format: Format,
    pub(crate) datetime_format: Format,
    pub(crate) null_string: Option<String>,
    pub(crate) table: Table,
    pub(crate) zoom: u16,
    pub(crate) screen_gridlines: bool,
    pub(crate) freeze_cell: (u32, u16),
    pub(crate) top_cell: (u32, u16),
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
            time_format: "hh:mm:ss;@".into(),
            date_format: "yyyy\\-mm\\-dd;@".into(),
            datetime_format: "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
            null_string: None,
            float_format: Format::default(),
            table: Table::new(),
            zoom: 100,
            screen_gridlines: true,
            freeze_cell: (0, 0),
            top_cell: (0, 0),
        }
    }
}
