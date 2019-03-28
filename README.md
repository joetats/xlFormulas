# xlFormulas
Helper class to write Excel-style formula strings to worksheets when saving from a Pandas dataframe.

Default initialization assumes the worksheet will be saved with an index and header row (the first real 'data' cell would be B2) but an index and header parameter are available to ensure alignment.

Pass in mathematical operators with strings, limited support currently for Excel built-in functions. If a value is not a column name in df.columns it is passed in as it is, whether that means it's an operator or builtin function.

The ```.formula()``` method returns a list of strings beginning with '=' and containing the row index for the Excel formula

Installation:

```pip install xl-formulas```

Basic usage:

```
import pandas as pd
from xlFormulas import ExcelFormulas

df = pd.read_excel('sample_data.xlsx')

# Pass in Pandas dataframe to intialize ExcelFormulas helper
ef = ExcelFormulas(df)

# Returns a column like "=B2+C2" in df['C']
df['C'] = ef.formula('A + B')

# Makes a "=(B2 + C2)/(C2 - D2)" column in df['D']
df['D'] = ef.formula(f'{ef.paren('A + B')} / {ef.paren('B - C')}'

# Use Excel built-in functions (Still pretty buggy)
# This would get a column of "=SUM(B2,C2,5)" in df['E']
df['E'] = ef.formula(ef.builtin('SUM', 'A', 'B', 'C'))
```