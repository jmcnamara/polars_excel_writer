# Sample Python Polars code to compare with the `perf_test.rs` example. The
# following is based on code from Alexander Beedie in the following GitHub
# thread: https://github.com/pola-rs/polars/issues/5568#issuecomment-1526316286

from codetiming import Timer
from datetime import date
import polars as pl

# Quickly spin-up a 1,000,000 element DataFrame.
df = pl.DataFrame(
    {
        "Int": range(250_000),
        "Float": 123.456789,
        "Date": date.today(),
        "String": "Test"
    }
)

# Export to Excel from polars.
with Timer():
    df.write_excel("dataframe_pl.xlsx")

# Export to Excel from pandas.
pf = df.to_pandas()
with Timer():
    pf.to_excel("dataframe_pd.xlsx", index=False)
