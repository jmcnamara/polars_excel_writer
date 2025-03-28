# polars_excel_writer

The `polars_excel_writer` crate is a library for serializing Polars dataframes
to Excel Xlsx files.

The crate uses [`rust_xlsxwriter`] to do the Excel serialization and is
typically 5x faster than Polars when exporting large dataframes to Excel.

It provides a primary interface [`PolarsXlsxWriter`] which is a configurable
Excel serializer that resembles the interface options provided by the Polars
[`write_excel()`] dataframe method.

The crate also provides a secondary [`ExcelWriter`] interface which is a simpler
Excel serializer that implements the Polars [`SerWriter`] trait to write a
dataframe to an Excel Xlsx file. However, unless you have existing code that
uses the [`SerWriter`] trait you should use the [`PolarsXlsxWriter`] interface.


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


## Performance

The table below shows the performance of writing a dataframe using Python
Polars, Python Pandas and `PolarsXlsxWriter`.

  | Test Case                     | Time (s) | Relative (%) |
  | :---------------------------- | :------- | :----------- |
  | `Polars`                      |     6.49 |         100% |
  | `Pandas`                      |    10.92 |         168% |
  | `polars_excel_writer`         |     1.22 |          19% |
  | `polars_excel_writer` + `zlib`|     1.08 |          17% |

See the [Performance] section of the docs for more detail.

[Performance]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/#performance

## See also

- [The `polars_excel_writer` crate].
- [The `polars_excel_writer` API docs at docs.rs].
- [The `polars_excel_writer` repository].
- [Roadmap of planned features].

[The `polars_excel_writer` crate]: https://crates.io/crates/polars_excel_writer
[The `polars_excel_writer` API docs at docs.rs]: https://docs.rs/polars_excel_writer/latest/polars_excel_writer/
[The `polars_excel_writer` repository]: https://github.com/jmcnamara/polars_excel_writer
[Release Notes and Changelog]: https://github.com/jmcnamara/polars_excel_writer/blob/main/CHANGELOG.md
[Roadmap of planned features]: https://github.com/jmcnamara/polars_excel_writer/issues/1