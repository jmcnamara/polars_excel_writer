/*!

# Comparison of `PolarsExcelWriter` and `Polars` `write_excel()`.

The `PolarsExcelWriter` crate is designed to be a drop-in replacement for the
Polars `write_excel()` function, providing a similar interface while leveraging
the performance benefits and native Rust support of the `rust_xlsxwriter`
library.

This comparison highlights the key similarities and differences between
`PolarsExcelWriter` and the Polars `write_excel()` function.

# Contents

- [Comparison of `PolarsExcelWriter` and `Polars` `write_excel()`.](#comparison-of-polarsexcelwriter-and-polars-write_excel)
- [Contents](#contents)
- [Background](#background)
- [Polars `write_excel()` parameters](#polars-write_excel-parameters)
  - [`workbook`](#workbook)
  - [`worksheet`](#worksheet)
  - [`position`](#position)
  - [`table_style`](#table_style)
  - [`table_name`](#table_name)
  - [`column_formats`](#column_formats)
  - [`dtype_formats`](#dtype_formats)
  - [`conditional_formats`](#conditional_formats)
  - [`header_format`](#header_format)
  - [`column_totals`](#column_totals)
  - [`column_widths`](#column_widths)
  - [`row_totals`](#row_totals)
  - [`row_heights`](#row_heights)
  - [`sparklines`](#sparklines)
  - [`formulas`](#formulas)
  - [`float_precision`](#float_precision)
  - [`include_header`](#include_header)
  - [`autofilter`](#autofilter)
  - [`autofit`](#autofit)
  - [`hidden_columns`](#hidden_columns)
  - [`hide_gridlines`](#hide_gridlines)
  - [`sheet_zoom`](#sheet_zoom)
  - [`freeze_panes`](#freeze_panes)



# Background

The Polars [`write_excel()`] function is maintained by the Polars core
developers and uses the Python [`XlsxWriter`] library to convert dataframes to
tables in Excel XLSX worksheets.

By comparison, `PolarsExcelWriter` uses the [`rust_xlsxwriter`] crate which is a
sister library to `XlsxWriter`. These XLSX libraries are written by the same
author and they have the same features and functionality.

Polars [`write_excel()`] currently has more functionality and better ease of
use, but similar functionality is being added to `PolarsExcelWriter`. See the
sections below.


[`XlsxWriter`]: https://xlsxwriter.readthedocs.io/index.html
[`write_excel()`]: https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel
[`rust_xlsxwriter`]: ../../rust_xlsxwriter/index.html

[Interacting with `rust_xlsxwriter`]: crate::excel_writer#interacting-with-rust_xlsxwriter

[Number Format Categories]: ../../rust_xlsxwriter/struct.Format.html#number-format-categories
[Number Formats in different locales]:  ../../rust_xlsxwriter/struct.Format.html#number-formats-in-different-locales

[Working with Conditional Formats]: ../../rust_xlsxwriter/conditional_format/index.html


# Polars `write_excel()` parameters

The following sections show the Polars `write_excel()` parameters and the
equivalent `PolarsExcelWriter` APIs.

## `workbook`

The `workbook` parameter is described in the Polars `write_excel()`
documentation as:

> `workbook : {str, Workbook}`
>
> String name or path of the workbook to create, BytesIO object, file opened
> in binary-mode, or an `xlsxwriter.Workbook` object that has not been closed.
> If None, writes to a `dataframe.xlsx` workbook in the working directory.

This functionality is implemented in `PolarsExcelWriter` using the following
APIs:

- [`PolarsExcelWriter::save()`] to save the dataframe to an XLSX file as a named
  `&str` or as a [`std::path`] `Path` or `PathBuf` instance. The option of using
  `None` to write to a default file called `dataframe.xlsx` isn't supported.

- [`PolarsExcelWriter::save_to_buffer()`]
  to save the dataframe to an XLSX file as a byte vector buffer.

- [`PolarsExcelWriter::save_to_writer()`] to save the dataframe to an XLSX file
  to a type that implements the `Write` trait.

An example of writing a Polars Rust dataframe to an Excel file.

```
# // This code is available in examples/doc_write_excel_write_dataframe.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() ->  PolarsResult<()>  {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Data" =>  &[10, 20, 15, 25, 30, 20],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_write_dataframe.png">

Unlike the Python Polars version it isn't possible to pass a `rust_xlsxwriter`
[`Workbook`] object directly to the API, however it is possible to create one
separately and use it to create a [`Worksheet`] object which can be passed. See
[Interacting with `rust_xlsxwriter`] and the example in the next section.



## `worksheet`

The `worksheet` parameter is described in the Polars `write_excel()`
documentation as:

> `worksheet : {str, Worksheet}`
>
> Name of target worksheet or an `xlsxwriter.Worksheet` object (in which
> case `workbook` must be the parent `xlsxwriter.Workbook` object); if None,
> writes to "Sheet1" when creating a new workbook (note that writing to an
> existing workbook requires a valid existing -or new- worksheet name).


The output worksheet name can be set using the
[`PolarsExcelWriter::set_worksheet_name()`] API. Otherwise, `PolarsExcelWriter`
will use the default Excel name of `Sheet1` (or `Sheet2`, `Sheet3`, etc. if more
than one worksheet is added).

An example setting the name for the output worksheet:
```
# // This code is available in examples/doc_write_excel_set_worksheet_name.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the worksheet name.
    excel_writer.set_worksheet_name("Polars Data")?;

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_worksheet_name.png">

The [`PolarsExcelWriter::worksheet()`] method can be used to get a reference to
the current/last worksheet in the workbook in order to manipulate it with a
`rust_xlsxwriter` [`Worksheet`] method.  For example the following code
demonstrates getting a reference to the worksheet used to write the dataframe
and setting its tab color.

```
# // This code is available in examples/doc_write_excel_worksheet.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Get the worksheet that the dataframe will be written to.
    let worksheet = excel_writer.worksheet()?;

    // Set the tab color for the worksheet using a `rust_xlsxwriter` worksheet
    // method.
    worksheet.set_tab_color("#FF9900");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_worksheet.png">

It is also possible to create a [`Worksheet`] separately and then write
the Polars dataframe to it using the
[`PolarsExcelWriter::write_dataframe_to_worksheet()`]
method. The latter is more useful if you need to do a lot of manipulation of
the worksheet or if you need to access some functionality not yet available in `polars_excel_writer`

Here is an example of creating and using a `rust_xlsxwriter` [`Worksheet`]
and then using it to write a Polars dataframe via
[`PolarsExcelWriter::write_dataframe_to_worksheet()`]. The worksheet instance is
then used to add a chart.

```
# // This code is available in examples/doc_write_excel_chart.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;
use rust_xlsxwriter::{Chart, ChartType, Workbook};

fn main() -> PolarsResult<()> {
    // Create a sample dataframe using Polars.
    let df: DataFrame = df!(
        "Data" => &[10, 20, 15, 25, 30, 20],
    )?;

    // Get some dataframe dimensions that we will use for the chart range.
    let row_min = 1; // Skip the header row.
    let row_max = df.height() as u32;

    // Create a new workbook and worksheet using rust_xlsxwriter.
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    // Write the dataframe to the worksheet using PolarsExcelWriter.
    let mut excel_writer = PolarsExcelWriter::new();
    excel_writer.write_dataframe_to_worksheet(&df, worksheet, 0, 0)?;

    // Move back to rust_xlsxwriter to create a new chart and have it plot the
    // range of the dataframe in the worksheet.
    let mut chart = Chart::new(ChartType::Line);
    chart
        .add_series()
        .set_values(("Sheet1", row_min, 0, row_max, 0));

    // Add the chart to the worksheet.
    worksheet.insert_chart(0, 2, &chart)?;

    // Save the file to disk.
    workbook.save("chart.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_chart.png">


## `position`

The `position` parameter is described in the Polars `write_excel()` documentation as:

> `position : {str, tuple}`
>
> Table position in Excel notation (eg: "A1"), or a (row,col) integer tuple.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::write_dataframe_to_cell()`] API to position a dataframe at
a specific `(row,col)` cell within a worksheet.

Using this method it is possible to write more than one dataframe to the
same worksheet, at different positions and without overlapping. For example:

```
# // This code is available in examples/doc_write_excel_write_dataframe_to_cell.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df1: DataFrame = df!(
        "Data 1" => &[10, 20, 15, 25, 30, 20],
    )?;

    let df2: DataFrame = df!(
        "Data 2" => &[1.23, 2.34, 3.56],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Write two dataframes to the same worksheet.
    excel_writer.write_dataframe_to_cell(&df1, 0, 0)?;
    excel_writer.write_dataframe_to_cell(&df2, 0, 2)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_write_dataframe_to_cell.png">


## `table_style`

The `table_style` parameter is described in the Polars `write_excel()` documentation as:

> `table_style : {str, dict}`
>
> A named Excel table style, such as "Table Style Medium 4", or a dictionary
> of `{"key":value,}` options containing one or more of the following keys:
> "style", "first_column", "last_column", "banded_columns, "banded_rows".

This parameter isn't currently implemented but it is on the backlog. The same
effect can be obtained using the [`PolarsExcelWriter::set_table()`] method and a
pre-configured `rust_xlsxwriter` [`Table`].


## `table_name`

The `table_name` parameter is described in the Polars `write_excel()` documentation as:

> `table_name : str`
>
> Name of the output table object in the worksheet; can then be referred to
> in the sheet by formulae/charts, or by subsequent `xlsxwriter` operations.

This parameter isn't currently implemented but it is on the backlog. The same
effect can be obtained using the [`PolarsExcelWriter::set_table()`] method and a
pre-configured `rust_xlsxwriter` [`Table`].


## `column_formats`

The `column_formats` parameter is described in the Polars `write_excel()` documentation as:

> `column_formats : dict`
>
> A `{colname(s):str,}` or `{selector:str,}` dictionary for applying an
> Excel format string to the given columns. Formats defined here (such as
> "dd/mm/yyyy", "0.00%", etc) will override any defined in `dtype_formats`.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_column_format()`] API for each required column.

The format can be a simple Excel number format string like `"$#,##0.00"` or a
more comprehensive `rust_xlsxwriter` [`Format`] that can have properties like
size, font, bold, italic or color.

Here is an example that demonstrates setting formats for different columns.

```
# // This code is available in examples/doc_write_excel_set_column_format.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "East" => &[1.0, 2.22, 3.333, 4.4444],
        "West" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the number formats for the columns.
    excel_writer.set_column_format("East", "0.00");
    excel_writer.set_column_format("West", "0.0000");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_column_format.png">


## `dtype_formats`

The `dtype_formats` parameter is described in the Polars `write_excel()` documentation as:

> `dtype_formats : dict`
>
> A `{dtype:str,}` dictionary that sets the default Excel format for the
> given dtype. (This can be overridden on a per-column basis by the
> `column_formats` param).

This is implemented in `PolarsExcelWriter` using the following APIs:

- [`PolarsExcelWriter::set_dtype_format()`] - for all supported data types, as shown below.
- [`PolarsExcelWriter::set_dtype_int_format()`] - for integer like data types.
- [`PolarsExcelWriter::set_dtype_float_format()`] - for float like data types.
- [`PolarsExcelWriter::set_dtype_number_format()`] - for number like data types (integers and floats).
- [`PolarsExcelWriter::set_dtype_datetime_format()`] - for datetime types.


The Polars' data types supported are:

- [`DataType::Boolean`]
- [`DataType::Int8`]
- [`DataType::Int16`]
- [`DataType::Int32`]
- [`DataType::Int64`]
- [`DataType::UInt8`]
- [`DataType::UInt16`]
- [`DataType::UInt32`]
- [`DataType::UInt64`]
- [`DataType::Float32`]
- [`DataType::Float64`]
- [`DataType::Date`]
- [`DataType::Time`]
- [`DataType::Datetime`]
- [`DataType::String`]
- [`DataType::Null`]

Here is an example that shows how to change Excel number format for floats.

```
# // This code is available in examples/doc_write_excel_float_format.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Float" => &[1000.0, 2000.22, 3000.333, 4000.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the float format.
    excel_writer.set_dtype_float_format("#,##0.00");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_float_format.png">


Here is another example that shows how to change the default format for Polars
time types.

```
# // This code is available in examples/doc_write_excel_time_format.rs
#
use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 00, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 18, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 37, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
        ],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the time format.
    excel_writer.set_dtype_format(DataType::Time, "hh:mm");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_time_format.png">

Here is another example that shows how to change the default format for Polars
date types.

```
# // This code is available in examples/doc_write_excel_date_format.rs
#
use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Date" => &[
            NaiveDate::from_ymd_opt(2023, 1, 11),
            NaiveDate::from_ymd_opt(2023, 1, 12),
            NaiveDate::from_ymd_opt(2023, 1, 13),
            NaiveDate::from_ymd_opt(2023, 1, 14),
        ],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the date format.
    excel_writer.set_dtype_format(DataType::Date, "mmm d yyyy");

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_date_format.png">




## `conditional_formats`

The `conditional_formats` parameter is described in the Polars `write_excel()` documentation as:

> `conditional_formats : dict`
>
> A dictionary of colname (or selector) keys to a format str, dict, or list
> that defines conditional formatting options for the specified columns.
>
> * If supplying a string typename, should be one of the valid `xlsxwriter`
>   types such as "3_color_scale", "data_bar", etc.
> * If supplying a dictionary you can make use of any/all `xlsxwriter`
>   supported options, including icon sets, formulae, etc.
> * Supplying multiple columns as a tuple/key will apply a single format
>   across all columns - this is effective in creating a heatmap, as the
>   min/max values will be determined across the entire range, not per-column.
> * Finally, you can also supply a list made up from the above options
>   in order to apply *more* than one conditional format to the same range.

This parameter isn't currently implemented but it is on the backlog.

However, it is possible to access conditional formatting by using the
`rust_xlsxwriter` APIs for the worksheet. See [Working with Conditional
Formats].


## `header_format`

The `header_format` parameter is described in the Polars `write_excel()` documentation as:

> `header_format : dict`
>
> A `{key:value,}` dictionary of `xlsxwriter` format options to apply
> to the table header row, such as `{"bold":True, "font_color":"#702963"}`.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_header()`] API. The format must be a `rust_xlsxwriter`
[`Format`] object.

This following example demonstrates setting the format for the header row.

```
# // This code is available in examples/doc_write_excel_set_header_format.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;
use rust_xlsxwriter::Format;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "East" => &[1, 1, 1, 1],
        "West" => &[2, 2, 2, 2],
        "North" => &[3, 3, 3, 3],
        "South" => &[4, 4, 4, 4],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Create and set the header format.
    let header_format = Format::new()
        .set_background_color("#C6EFCE")
        .set_font_color("#006100")
        .set_bold();

    // Set the number formats for the columns.
    excel_writer.set_header_format(&header_format);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_header_format.png">



## `column_totals`

The `column_totals` parameter is described in the Polars `write_excel()` documentation as:

> `column_totals : {bool, list, dict}`
>
> Add a column-total row to the exported table.
>
> * If True, all numeric columns will have an associated total using "sum".
> * If passing a string, it must be one of the valid total function names
>   and all numeric columns will have an associated total using that function.
> * If passing a list of colnames, only those given will have a total.
> * For more control, pass a `{colname:funcname,}` dict.
>
> Valid column-total function names are "average", "count_nums", "count",
> "max", "min", "std_dev", "sum", and "var".

This parameter isn't currently implemented but it is on the backlog. The same
effect can be obtained using the [`PolarsExcelWriter::set_table()`] method and a
pre-configured `rust_xlsxwriter` [`Table`]. See also `rust_xlsxwriter`
[`TableColumn`] and [`TableColumn::set_total_function()`].


## `column_widths`

The `column_widths` parameter is described in the Polars `write_excel()` documentation as:

> `column_widths : {dict, int}`
>
> A `{colname:int,}` or `{selector:int,}` dict or a single integer that
> sets (or overrides if autofitting) table column widths, in integer pixel
> units. If given as an integer the same value is used for all table columns.

This parameter isn't currently implemented but it is on the backlog. The same
effect can be achieved using the `rust_xlsxwriter` [`Worksheet`] object and the
[`Worksheet::set_column_width_pixels()`] method.


## `row_totals`

The `row_totals` parameter is described in the Polars `write_excel()` documentation as:

> `row_totals : {dict, list, bool}`
>
> Add a row-total column to the right-hand side of the exported table.
>
> * If True, a column called "total" will be added at the end of the table
>   that applies a "sum" function row-wise across all numeric columns.
> * If passing a list/sequence of column names, only the matching columns
>   will participate in the sum.
> * Can also pass a `{colname:columns,}` dictionary to create one or
>   more total columns with distinct names, referencing different columns.

This parameter isn't currently implemented but it is on the backlog.


## `row_heights`

The `row_heights` parameter is described in the Polars `write_excel()` documentation as:

> `row_heights : {dict, int}`
>
> An int or `{row_index:int,}` dictionary that sets the height of the given
> rows (if providing a dictionary) or all rows (if providing an integer) that
> intersect with the table body (including any header and total row) in
> integer pixel units. Note that `row_index` starts at zero and will be
> the header row (unless `include_header` is False).

This parameter isn't currently implemented but it is on the backlog. The same
effect can be achieved using the `rust_xlsxwriter` [`Worksheet`] object and the
[`Worksheet::set_row_height_pixels()`] method.


## `sparklines`

The `sparklines` parameter is described in the Polars `write_excel()` documentation as:

> `sparklines : dict`
>
> A `{colname:list,}` or `{colname:dict,}` dictionary defining one or more
> sparklines to be written into a new column in the table.
>
> * If passing a list of colnames (used as the source of the sparkline data)
>   the default sparkline settings are used (eg: line chart with no markers).
> * For more control an `xlsxwriter`-compliant options dict can be supplied,
>   in which case three additional polars-specific keys are available:
>   "columns", "insert_before", and "insert_after". These allow you to define
>   the source columns and position the sparkline(s) with respect to other
>   table columns. If no position directive is given, sparklines are added to
>   the end of the table (eg: to the far right) in the order they are given.

This parameter isn't currently implemented but it is on the backlog. The same
effect can be achieved using the `rust_xlsxwriter` [`Worksheet`] object and the
[`Sparkline`] object.

## `formulas`

The `formulas` parameter is described in the Polars `write_excel()` documentation as:

> `formulas : dict`
>
> A `{colname:formula,}` or `{colname:dict,}` dictionary defining one or
> more formulas to be written into a new column in the table. Note that you
> are strongly advised to use structured references in your formulae wherever
> possible to make it simple to reference columns by name.
>
> * If providing a string formula (such as `=[@colx]*[@coly]`) the column will
>   be added to the end of the table (eg: to the far right), after any default
>   sparklines and before any row_totals.
> * For the most control supply an options dictionary with the following keys:
>   "formula" (mandatory), one of "insert_before" or "insert_after", and
>   optionally "return_dtype". The latter is used to appropriately format the
>   output of the formula and allow it to participate in row/column totals.

This parameter isn't currently implemented but it is on the backlog.


## `float_precision`

The `float_precision` parameter is described in the Polars `write_excel()` documentation as:

> `float_precision : int`
>
> Default number of decimals displayed for floating point columns (note that
> this is purely a formatting directive; the actual values are not rounded).

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_float_precision()`] API to set the number precision of
all floats exported from the dataframe to Excel. The precision is converted to
an Excel number format (using the
[`PolarsExcelWriter::set_dtype_float_format()`] method).

Note, the numeric values aren't truncated in Excel, this option just
controls the display of the number.

The following example demonstrates how to set the precision of the float output.
Setting the precision to 3 is equivalent to an Excel number format of `0.000`.

```
# // This code is available in examples/doc_write_excel_float_precision.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the float precision.
    excel_writer.set_float_precision(3);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_float_precision.png">


## `include_header`

The `include_header` parameter is described in the Polars `write_excel()` documentation as:

> `include_header : bool`
>
> Indicate if the table should be created with a header row.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_header()`] API which turns off the table header row:

```
# // This code is available in examples/doc_write_excel_set_header.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Int" => &[1, 2, 3, 4],
        "Float" => &[1.0, 2.22, 3.333, 4.4444],
    )?;

    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Turn off the default header.
    excel_writer.set_header(false);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_has_header_off.png">



## `autofilter`

The `autofilter` parameter is described in the Polars `write_excel()` documentation as:

> `autofilter : bool`
>
> If the table has headers, provide autofilter capability.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_autofilter()`] API to turn on/off the autofilter in the
dataframe table. It is on by default.


## `autofit`

The`autofit` parameter is described in the Polars `write_excel()` documentation as:

> `autofit : bool`
>
> Calculate individual column widths from the data.

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_autofit()`] API to adjust the dataframe columns to the
maximum data width in each column:

```
# // This code is available in examples/doc_write_excel_autofit.rs
#
use polars::prelude::*;

use polars_excel_writer::PolarsExcelWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "Col 1" => &["A", "B", "C", "D"],
        "Col 2" => &["Hello", "World", "Hello, world", "Ciao"],
        "Col 3" => &[1.234578, 123.45678, 123456.78, 12345679.0],
    )?;

    // Create a new Excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set an number format for column 3.
    excel_writer.set_column_format("Col 3", "$#,##0.00");

    // Autofit the output data.
    excel_writer.set_autofit(true);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

<img src="https://rustxlsxwriter.github.io/images/excelwriter_autofit.png">


## `hidden_columns`

The `hidden_columns` parameter is described in the Polars `write_excel()` documentation as:

> `hidden_columns : str | list`
>
>  A column name, list of column names, or a selector representing table
>  columns to mark as hidden in the output worksheet.

This parameter isn't currently implemented but it is on the backlog. The same
effect can be achieved using the `rust_xlsxwriter` [`Worksheet`] object and the
[`Worksheet::set_column_hidden()`] method.


## `hide_gridlines`

The `hide_gridlines` parameter is described in the Polars `write_excel()` documentation as:

> `hide_gridlines : bool`
>
> Do not display any gridlines on the output worksheet.

This is implemented in `PolarsExcelWriter` using the [`PolarsExcelWriter::set_screen_gridlines()`] API:

```
# // This code is available in examples/doc_write_excel_set_screen_gridlines.rs
#
# use polars::prelude::*;
#
# use polars_excel_writer::PolarsExcelWriter;
#
# fn main() -> PolarsResult<()> {
#     // Create a sample dataframe for the example.
#     let df: DataFrame = df!(
#         "String" => &["North", "South", "East", "West"],
#         "Int" => &[1, 2, 3, 4],
#         "Float" => &[1.0, 2.22, 3.333, 4.4444],
#     )?;
#
    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Turn off the screen gridlines.
    excel_writer.set_screen_gridlines(false);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;
#
#     Ok(())
# }
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_screen_gridlines.png">


## `sheet_zoom`

The `sheet_zoom` parameter is described in the Polars `write_excel()` documentation as:

> `sheet_zoom : int`
>
> Set the default zoom level of the output worksheet.

This is implemented in `PolarsExcelWriter` using the [`PolarsExcelWriter::set_zoom()`] API:

```
# // This code is available in examples/doc_write_excel_set_zoom.rs
#
# use polars::prelude::*;
#
# use polars_excel_writer::PolarsExcelWriter;
#
# fn main() -> PolarsResult<()> {
#     // Create a sample dataframe for the example.
#     let df: DataFrame = df!(
#         "String" => &["North", "South", "East", "West"],
#         "Int" => &[1, 2, 3, 4],
#         "Float" => &[1.0, 2.22, 3.333, 4.4444],
#     )?;
#
    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Set the worksheet zoom level.
    excel_writer.set_zoom(200);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;
#
#     Ok(())
# }
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_zoom.png">


## `freeze_panes`

The `freeze_panes` parameter is described in the Polars `write_excel()` documentation as:

> `freeze_panes : str | (str, int, int) | (int, int) | (int, int, int, int)`
>
> Freeze workbook panes.
>
> * If (row, col) is supplied, panes are split at the top-left corner of the
>   specified cell, which are 0-indexed. Thus, to freeze only the top row,
>   supply (1, 0).
> * Alternatively, cell notation can be used to supply the cell. For example,
>   "A2" indicates the split occurs at the top-left of cell A2, which is the
>   equivalent of (1, 0).
> * If (row, col, top_row, top_col) are supplied, the panes are split based on
>   the `row` and `col`, and the scrolling region is initialized to begin at
>   the `top_row` and `top_col`. Thus, to freeze only the top row and have the
>   scrolling region begin at row 10, column D (5th col), supply (1, 0, 9, 4).
>   Using cell notation for (row, col), supplying ("A2", 9, 4) is equivalent.
>

This is implemented in `PolarsExcelWriter` using the
[`PolarsExcelWriter::set_freeze_panes()`] API.

This example demonstrates freezing the top row of the dataframe:

```
# // This code is available in examples/doc_write_excel_set_freeze_panes.rs
#
# use polars::prelude::*;
#
# use polars_excel_writer::PolarsExcelWriter;
#
# fn main() -> PolarsResult<()> {
#     // Create a sample dataframe for the example.
#     let df: DataFrame = df!(
#         "String" => &["North", "South", "East", "West"],
#         "Int" => &[1, 2, 3, 4],
#         "Float" => &[1.0, 2.22, 3.333, 4.4444],
#     )?;
#
    // Create a new excel writer.
    let mut excel_writer = PolarsExcelWriter::new();

    // Freeze the top row.
    excel_writer.set_freeze_panes(1, 0);

    // Write the dataframe to Excel.
    excel_writer.write_dataframe(&df)?;

    // Save the file to disk.
    excel_writer.save("dataframe.xlsx")?;
#
#     Ok(())
# }
```

<img src="https://rustxlsxwriter.github.io/images/write_excel_set_freeze_panes.png">

*/

// Imports to get access to inter-documentation links.
#[allow(unused_imports)]
use crate::excel_writer::PolarsExcelWriter;
#[allow(unused_imports)]
use polars::prelude::*;
#[allow(unused_imports)]
use rust_xlsxwriter::{Format, Sparkline, Table, TableColumn, Workbook, Worksheet};
