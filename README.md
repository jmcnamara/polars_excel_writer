# polars_excel_writer

The `polars_excel_writer` crate is a library for serializing Polars dataframes
to Excel Xlsx files.

It provides a primary interface [`PolarsXlsxWriter`] which is a configurable
Excel serializer that resembles the interface options provided by the Polars
[`write_excel()`] dataframe method.

It also provides a secondary [`ExcelWriter`] interface which is a simpler
Excel serializer that implements the Polars [`SerWriter`] trait to write a
dataframe to an Excel Xlsx file. However, unless you have existing code that
uses the [`SerWriter`] trait you should use the [`PolarsXlsxWriter`]
interface.

Unless you have existing code that uses the Polars [`SerWriter`] trait you
should use the primary [`PolarsXlsxWriter`] interface.

This crate uses [`rust_xlsxwriter`] to do the Excel serialization.

[`ExcelWriter`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/write/struct.ExcelWriter.html
[`PolarsXlsxWriter`]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/xlsx_writer/struct.PolarsXlsxWriter.html

[`SerWriter`]:
    https://docs.rs/polars/latest/polars/prelude/trait.SerWriter.html

[`rust_xlsxwriter`]: https://docs.rs/rust_xlsxwriter/latest/rust_xlsxwriter/

[`write_excel()`]:
   https://pola-rs.github.io/polars/py-polars/html/reference/api/polars.DataFrame.write_excel.html#polars.DataFrame.write_excel

## Example

An example of writing a Polar Rust dataframe to an Excel file using the
`PolarsXlsxWriter` interface.

```rust
use chrono::prelude::*;
use polars::prelude::*;

use polars_excel_writer::PolarsXlsxWriter;

fn main() -> PolarsResult<()> {
    // Create a sample dataframe for the example.
    let df: DataFrame = df!(
        "String" => &["North", "South", "East", "West"],
        "Integer" => &[1, 2, 3, 4],
        "Float" => &[4.0, 5.0, 6.0, 7.0],
        "Time" => &[
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            NaiveTime::from_hms_milli_opt(2, 59, 3, 456).unwrap(),
            ],
        "Date" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap(),
            ],
        "Datetime" => &[
            NaiveDate::from_ymd_opt(2022, 1, 1).unwrap().and_hms_opt(1, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 2).unwrap().and_hms_opt(2, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 3).unwrap().and_hms_opt(3, 0, 0).unwrap(),
            NaiveDate::from_ymd_opt(2022, 1, 4).unwrap().and_hms_opt(4, 0, 0).unwrap(),
        ],
    )?;

    // Create a new Excel writer.
    let mut xlsx_writer = PolarsXlsxWriter::new();

    // Write the dataframe to Excel.
    xlsx_writer.write_dataframe(&df)?;

    // Save the file to disk.
    xlsx_writer.save("dataframe.xlsx")?;

    Ok(())
}
```

Output file:

<img src="https://rustxlsxwriter.github.io/images/write_excel_combined.png">

## See also

- [Changelog]: Recent additions and fixes.
- [Performance]: Performance comparison with Python based methods.

[Changelog]: https://github.com/jmcnamara/polars_excel_writer/blob/main/CHANGELOG.md
[Performance]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/#performance
