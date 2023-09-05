# Changelog

All notable changes to `polars_excel_writer` will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).


## [0.3.0] - 2023-09-05

### Added

More worksheet utility methods.

- Added support for renaming worksheets via the [`set_worksheet_name()`] method.

- Added support for adding worksheets via the [`add_worksheet`] method. This
  allows you to add dataframes to several different worksheets in a workbook.

- Added support for accessing the current worksheets via the [`worksheet()`] method.

[`set_worksheet_name()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/xlsx_writer/struct.PolarsXlsxWriter.html#method.set_worksheet_name

[`add_worksheet()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/xlsx_writer/struct.PolarsXlsxWriter.html#method.add_worksheet

[`worksheet()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/xlsx_writer/struct.PolarsXlsxWriter.html#method.worksheet


## [0.2.0] - 2023-09-04

### Added

- Added support for setting worksheet table properties via the PolarsXlsxWriter
  [`set_table()`] method.

[`set_table()`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/xlsx_writer/struct.PolarsXlsxWriter.html#method.set_table

## [0.1.0] - 2023-08-20

### Added

- First functional version.

