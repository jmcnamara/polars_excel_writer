// excel_writer - A Polars extension to serialize dataframes to Excel xlsx files.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2023, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

use std::path::Path;

use polars::export::arrow::temporal_conversions::{
    date32_to_date, time64ns_to_time, timestamp_ms_to_datetime, timestamp_ns_to_datetime,
    timestamp_us_to_datetime,
};
use polars::prelude::*;
use rust_xlsxwriter::{Format, Table, Workbook, Worksheet, XlsxError};

/// `PolarsXlsxWriter` provides an Excel XLSX serializer that works with Polars
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
///     let mut writer = PolarsXlsxWriter::new();
///
///     writer.write_dataframe(df)?;
///     writer.write_excel("dataframe.xlsx")?;
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
/// also plot the data on an Excel chart. We can use `PolarsXlsxWriter` for the
/// data writing part and `rust_xlsxwriter` for all the other functionality.
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
///     let mut writer = PolarsXlsxWriter::new();
///     writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
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
/// The `rust_xlsxwriter` crate uses an error type called [`XlsxError`] while
/// Polars and `PolarsXlsxWriter` use an error type called [`PolarsError`]. In
/// order to make interoperability with Polars easier the
/// `rust_xlsxwriter::XlsxError` type maps to (and from) the `PolarsError` type.
///
/// That is why in the previous example we were able to use two different error
/// types within the same result/error context of `PolarsError`. Note:
/// `PolarsResult<T>` expands to `Result<T, PolarsError>`.
///
/// In order for this to be enabled you must use the `rust_xlsxwriter` `polars`
/// crate feature. For convenience it is turned on automatically when you use
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
    /// [`write_excel()`](PolarsXlsxWriter::write_excel).
    /// # Errors
    ///
    /// A [`PolarsError`] `ComputeError` error that wraps any `rust_xlsxwriter`
    /// errors in a string.
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
    ///     let mut writer = PolarsXlsxWriter::new();
    ///
    ///     writer.write_dataframe(&df)?;
    ///     writer.write_excel("dataframe.xlsx")?;
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
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        Ok(())
    }

    /// Write a dataframe to a user defined cell in a worksheet.
    ///
    /// Writes the supplied dataframe to a user defined cell in the first sheet
    /// of a new Excel workbook.
    ///
    /// Since the dataframe can be positioned with the worksheet it is possible
    /// to write more than one to the same worksheet (without overlapping).
    ///
    /// The worksheet must be written to a file using
    /// [`write_excel()`](PolarsXlsxWriter::write_excel).
    ///
    /// # Errors
    ///
    /// A [`PolarsError`] `ComputeError` error that wraps any `rust_xlsxwriter`
    /// errors in a string.
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
    ///     let mut writer = PolarsXlsxWriter::new();
    ///
    ///     // Write two dataframes to the same worksheet.
    ///     writer.write_dataframe_to_cell(&df1, 0, 0)?;
    ///     writer.write_dataframe_to_cell(&df2, 0, 2)?;
    ///
    ///     writer.write_excel("dataframe.xlsx")?;
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
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, row, col, &options)?;

        Ok(())
    }

    /// Write a dataframe to a user supplied worksheet.
    ///
    /// Writes the dataframe to a `rust_xlsxwriter` [`Worksheet`] object. This
    /// worksheet cannot be saved via
    /// [`write_excel()`](PolarsXlsxWriter::write_excel). Instead it must be
    /// used in conjunction with a `rust_xlsxwriter` [`Workbook`].
    ///
    /// This is useful for mixing `PolarsXlsxWriter` data writing with
    /// additional Excel functionality provided by `rust_xlsxwriter`. See
    /// [Interacting with `rust_xlsxwriter`](#interacting-with-rust_xlsxwriter)
    /// and the example below.
    ///
    /// # Errors
    ///
    /// A [`PolarsError`] `ComputeError` error that wraps any `rust_xlsxwriter`
    /// errors in a string.
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
    ///     let mut writer = PolarsXlsxWriter::new();
    ///     writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;
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
    /// The `write_excel()` method writes all the workbook and worksheet data to
    /// a new xlsx file. It will overwrite any existing file.
    ///
    /// The method can be called multiple times so it is possible to get
    /// incremental files at different stages of a process, or to save the same
    /// Workbook object to different paths. However, `write_excel()` is an
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
    /// A [`PolarsError`] `ComputeError` error that wraps any `rust_xlsxwriter`
    /// errors in a string.
    ///
    pub fn write_excel<P: AsRef<Path>>(&mut self, path: P) -> PolarsResult<()> {
        self.workbook.save(path)?;

        Ok(())
    }

    /// Turn on/off the dataframe header in the exported Excel file.
    ///
    /// Turn on/off the dataframe header row in the Excel table. It is on by
    /// default.
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
    ///     let mut writer = PolarsXlsxWriter::new();
    ///
    ///     writer.set_header(false);
    ///
    ///     writer.write_dataframe(df)?;
    ///     writer.write_excel("dataframe.xlsx")?;
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
        self.options.has_header = has_header;
        self.options.table.set_header_row(has_header);
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
    /// TODO
    ///
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
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This example
    /// demonstrates how to change the default format for Polars date types.
    ///
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
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to change the default format for Polars
    /// datetime types.
    ///
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
    /// number format string. These format strings can be obtained for the
    /// Format Cells -> Number dialog in Excel.
    ///
    /// See all the [Number Format Categories] and subsequent sections in the
    /// `rust_xlsxwriter` documentation.
    ///
    /// [Number Format Categories]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Format.html#number-format-categories
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting an Excel number format for floats.
    ///
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
    /// The precision should be in the Excel range 1-30.
    ///
    /// Note, the numeric values aren't truncated in Excel, this option just
    /// controls the display of the number.
    ///
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates how to set the precision of the float output.
    /// Setting the precision to 3 is equivalent to an Excel number format of
    /// `0.000`.
    ///
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
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// demonstrates setting a value for Null values in the dataframe. The
    /// default is to write them as blank cells.
    ///
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
    /// There are several limitations to this autofit method, see the
    /// `rust_xlsxwriter` docs on [`worksheet.autofit()`] for details.
    ///
    /// [`worksheet.autofit()`]:
    ///     https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/struct.Worksheet.html#method.autofit
    ///
    /// TODO
    ///
    /// An example of writing a Polar Rust dataframe to an Excel file. This
    /// example demonstrates autofitting column widths in the output worksheet.
    ///
    ///
    /// Output file:
    ///
    /// <img
    /// src="https://rustxlsxwriter.github.io/images/excelwriter_autofit.png">
    ///
    /// TODO
    pub fn set_autofit(&mut self, use_autofit: bool) -> &mut PolarsXlsxWriter {
        self.options.use_autofit = use_autofit;
        self
    }
    // -----------------------------------------------------------------------
    // Internal functions/methods.
    // -----------------------------------------------------------------------

    // TODO
    pub(crate) fn write_to_buffer(&mut self, df: &DataFrame) -> PolarsResult<Vec<u8>> {
        let options = self.options.clone();
        let worksheet = self.last_worksheet()?;

        Self::write_dataframe_internal(df, worksheet, 0, 0, &options)?;

        let buf = self.workbook.save_to_buffer()?;

        Ok(buf)
    }

    fn last_worksheet(&mut self) -> PolarsResult<&mut Worksheet> {
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

    // TODO
    #[allow(clippy::too_many_lines)]
    fn write_dataframe_internal(
        df: &DataFrame,
        worksheet: &mut Worksheet,
        row_offset: u32,
        col_offset: u16,
        options: &WriterOptions,
    ) -> Result<(), XlsxError> {
        let header_offset = u32::from(options.has_header);

        // Iterate through the dataframe column by column.
        for (col_num, column) in df.get_columns().iter().enumerate() {
            let col_num = col_offset + col_num as u16;

            // Store the column names for use as table headers.
            if options.has_header {
                worksheet.write(row_offset, col_num, column.name())?;
            }

            // Write the row data for each column/type.
            for (row_num, data) in column.iter().enumerate() {
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
                    AnyValue::Utf8(value) => {
                        worksheet.write_string(row_num, col_num, value)?;
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
                            &datetime,
                            &options.datetime_format,
                        )?;
                        worksheet.set_column_width(col_num, 18)?;
                    }
                    AnyValue::Date(value) => {
                        let date = date32_to_date(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &date,
                            &options.date_format,
                        )?;
                        worksheet.set_column_width(col_num, 10)?;
                    }
                    AnyValue::Time(value) => {
                        let time = time64ns_to_time(value);
                        worksheet.write_datetime_with_format(
                            row_num,
                            col_num,
                            &time,
                            &options.time_format,
                        )?;
                    }
                    _ => {
                        eprintln!(
                            "WARNING: Polars AnyValue data type '{}' is not supported by Excel",
                            data.dtype()
                        );
                        break;
                    }
                }
            }
        }

        // Create a table for the dataframe range.
        let (mut max_row, max_col) = df.shape();
        if !options.has_header {
            max_row -= 1;
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

        Ok(())
    }
}

// -----------------------------------------------------------------------
// TODO
// -----------------------------------------------------------------------

/// TODO
#[derive(Clone)]
pub(crate) struct WriterOptions {
    pub(crate) has_header: bool,
    pub(crate) use_autofit: bool,
    pub(crate) date_format: Format,
    pub(crate) time_format: Format,
    pub(crate) float_format: Format,
    pub(crate) datetime_format: Format,
    pub(crate) null_string: Option<String>,
    pub(crate) table: Table,
}

impl Default for WriterOptions {
    fn default() -> Self {
        Self::new()
    }
}

// TODO
impl WriterOptions {
    // TODO
    fn new() -> WriterOptions {
        WriterOptions {
            has_header: true,
            use_autofit: false,
            time_format: "hh:mm:ss;@".into(),
            date_format: "yyyy\\-mm\\-dd;@".into(),
            datetime_format: "yyyy\\-mm\\-dd\\ hh:mm:ss".into(),
            null_string: None,
            float_format: Format::default(),
            table: Table::new(),
        }
    }
}
