# excelparser

Simple excel parser written in python, it's essentially a wrapper of Panda, openpyxl. Supports csv, xls, xlsx.

### Example

```py
import excelparser

ep = excelparser.ExcelParser('/some/dir/yourfile.xlsx')
ep.parse()

# convert column value
ep.cvt_col_name('amount', lambda x: 0 if x == "" else float(x) * 1000)

# read value at column 'amount'
for i in range(ep.row_count()):
    amt = ep.getcol('amount', i)
    print(f"row: {i}, amt: {amt}")

# make copy of the selected columns
ep = ep.copy_col_name(['product', 'amount'])

# save the copy
ep.export('/some/dir/newfile.xlsx')
```
