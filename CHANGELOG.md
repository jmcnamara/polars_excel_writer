# Changelog

All notable changes to `polars_excel_writer` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [0.15.0] - 2025-06-15

### Added

- Update dependencies to `rust_xlsxwriter` 0.88 and `polars` 0.48.

- This version emphasizes the `PolarsExcelWriter`interface and compatibility
  with the Polars `write_excel` interface. See the [`polars_excel_writer`
  Roadmap].

[`polars_excel_writer` Roadmap]: https://github.com/jmcnamara/polars_excel_writer/issues/1

### Deprecated

- This version of crate drops support for the Polars `SerWriter` interface in
  order to maximize compatibility with the Polars `write_excel` interface. The
  `SerWriter` interface will eventually move to another crate.

- Rename `PolarsXlsxWriter` struct to `PolarsExcelWriter`.

  This rename is for consistency with Polars. The deprecated `PolarsXlsxWriter`
  struct name is still supported via a type alias.


## [0.14.0] - 2025-05-03

### Added

- Added support for setting dataframe formatting based on data types or columns.
  It also adds header formatting. See:

  - `PolarsExcelWriter::set_dtype_format()`
  - `PolarsExcelWriter::set_column_format()`
  - `PolarsExcelWriter::set_header_format()`

### Deprecated

- The following functions are deprecated in favour of
  `PolarsExcelWriter::set_dtype_format()` and variants:

  - `PolarsExcelWriter::set_float_format()`
  - `PolarsExcelWriter::set_time_format()`
  - `PolarsExcelWriter::set_date_format()`
  - `PolarsExcelWriter::set_datetime_format()`


## [0.13.0] - 2025-03-15

### Added

- Update dependencies to `rust_xlsxwriter` 0.84.


## [0.12.0] - 2025-01-29

### Added

- Update dependencies to `rust_xlsxwriter` 0.82.0 and `polars` 0.46.

- Added support for overriding the default handling of NaN and Infinity numbers.
  These aren't supported by Excel so they are replaced with default or custom
  string values. See:

  - `PolarsExcelWriter::set_nan_value()`
  - `PolarsExcelWriter::set_infinity_value()`
  - `PolarsExcelWriter::set_neg_infinity_value()`


## [0.11.0] - 2025-01-18

### Added

- Update dependencies to `rust_xlsxwriter` 0.81.0 and `polars` 0.45.


## [0.10.0] - 2025-01-05

### Added

- Update dependencies to `rust_xlsxwriter` 0.80.0 and `polars` 0.44.

- Changed documentation to highlight `write_xlsx` as the primary interface,
  since that will be the main interface in future releases.


## [0.9.0] - 2024-09-18

### Added

- Update dependencies to `rust_xlsxwriter` 0.77.0 and `polars` 0.43.


## [0.8.0] - 2024-08-24

### Added

- Update dependencies to `rust_xlsxwriter` 0.74.0 and `polars` 0.42.0.


## [0.7.0] - 2024-03-25

### Added

- Update dependencies to `rust_xlsxwriter` 0.63.0 and `polars` 0.37.0.


## [0.6.0] - 2024-01-24

### Added

- Update dependencies to `rust_xlsxwriter` 0.62.0 and `polars` 0.36.2.


## [0.5.0] - 2024-01-15

### Added

- Added support for writing `u64` and `i64` number within Excel's limitations.
  This implies a loss of precision outside Excel's integer range of +/-
  999,999,999,999,999 (15 digits).


## [0.4.0] - 2023-11-22

### Added

- Update to the latest `rust_xlsxwriter` to fix issues with `PolarsError` type/location.


## [0.3.0] - 2023-09-05

### Added

More worksheet utility methods.

- Added support for renaming worksheets via the [`set_worksheet_name()`] method.

- Added support for adding worksheets via the [`add_worksheet()`] method. This
  allows you to add dataframes to several different worksheets in a workbook.

- Added support for accessing the current worksheets via the [`worksheet()`] method.

[`set_worksheet_name()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/excel_writer/struct.PolarsExcelWriter.html#method.set_worksheet_name

[`add_worksheet()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/excel_writer/struct.PolarsExcelWriter.html#method.add_worksheet

[`worksheet()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/excel_writer/struct.PolarsExcelWriter.html#method.worksheet


## [0.2.0] - 2023-09-04

### Added

- Added support for setting worksheet table properties via the `PolarsExcelWriter`
  [`set_table()`] method.

[`set_table()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/excel_writer/struct.PolarsExcelWriter.html#method.set_table

## [0.1.0] - 2023-08-20

### Added

- First functional version.

