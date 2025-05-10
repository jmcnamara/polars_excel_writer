# Examples for the `polars_excel_writer` library.

This directory contains working examples showing different features of the
`polars_excel_writer` library.

The `app_{name}.rs` examples are small complete programs showing a feature or
collection of features.

The `doc_{struct}_{function}.rs` examples are more specific examples from the
documentation and generally show how an individual function works.

* `doc_write_excel_add_worksheet.rs` - An example of writing a Polar Rust
  dataframes to separate worksheets in an Excel workbook.

* `doc_write_excel_autofit.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This example demonstrates autofitting column
  widths in the output worksheet.

* `doc_write_excel_chart.rs` - An example of using `polars_excel_writer` in
  conjunction with `rust_xlsxwriter` to write a Polars dataframe to a
  worksheet and then add a chart to plot the data.

* `doc_write_excel_combined.rs` - An example of writing a Polar Rust
  dataframe to an Excel file.

* `doc_write_excel_date_format.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This example demonstrates how to change the
  default format for Polars date types.

* `doc_write_excel_datetime_format.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This example demonstrates how to change the
  default format for Polars datetime types.

* `doc_write_excel_float_format.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates setting an Excel number
  format for floats.

* `doc_write_excel_float_precision.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This example demonstrates how to set the
  precision of the float output. Setting the precision to 3 is equivalent
  to an Excel number format of `0.000`.

* `doc_write_excel_intro.rs` - An example of writing a Polar Rust dataframe
  to an Excel file.

* `doc_write_excel_null_values.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates setting a value for Null
  values in the dataframe. The default is to write them as blank cells.

* `doc_write_excel_set_column_format.rs` - An example of writing a Polar
  Rust dataframe to an Excel file. This demonstrates setting formats for
  different columns.

* `doc_write_excel_set_freeze_panes_top_cell.rs` - An example of writing a
  Polar Rust dataframe to an Excel file. This demonstrates freezing the top
  row and setting a non-default first row within the pane.

* `doc_write_excel_set_freeze_panes.rs` - An example of writing a Polar
  Rust dataframe to an Excel file. This demonstrates freezing the top row.

* `doc_write_excel_set_header_format.rs` - An example of writing a Polar
  Rust dataframe to an Excel file. This demonstrates setting the format for
  the header row.

* `doc_write_excel_set_header.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates saving the dataframe
  without a header.

* `doc_write_excel_set_nan_value.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates handling NaN and Infinity
  values with custom string representations.

* `doc_write_excel_set_screen_gridlines.rs` - An example of writing a Polar
  Rust dataframe to an Excel file. This demonstrates turning off the screen
  gridlines.

* `doc_write_excel_set_table.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates setting properties of the
  worksheet table that wraps the output dataframe.

* `doc_write_excel_set_worksheet_name.rs` - An example of writing a Polar
  Rust dataframe to an Excel file. This demonstrates setting the name for
  the output worksheet.

* `doc_write_excel_set_zoom.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates setting the worksheet zoom
  level.

* `doc_write_excel_time_format.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This example demonstrates how to change the
  default format for Polars time types.

* `doc_write_excel_worksheet.rs` - An example of writing a Polar Rust
  dataframe to an Excel file. This demonstrates getting a reference to the
  worksheet used to write the dataframe and setting its tab color.

* `doc_write_excel_write_dataframe_to_cell.rs` - An example of writing more
  than one Polar dataframes to an Excel worksheet.

* `doc_write_excel_write_dataframe.rs` - An example of writing a Polar Rust
  dataframe to an Excel file.

* `perf_test.rs` - Simple performance test to compare with the Python
  Polars example in `perf_test.py`.

