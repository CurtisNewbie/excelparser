# excelparser

Simple excel parser written in python, it's essentially a wrapper of Panda, openpyxl. Supports csv, xls, xlsx.

To install it, simply run `python3 -m pip install .`.

e.g.,

```sh
python3 -m pip install . -i https://pypi.tuna.tsinghua.edu.cn/simple
```

Or you can use:

```sh
curl https://raw.githubusercontent.com/CurtisNewbie/excelparser/main/install.sh | bash
```

### Example

```py
import excelparser

# parse the xlsx file
ep = excelparser.parse('/some/dir/yourfile.xlsx')

# convert column value
ep.cvt_col_name('amount', lambda x: 0 if x == "" else float(x) * 1000)

# read value at column 'amount'
for i in range(ep.row_count()):
    amt = ep.get_col('amount', i)
    print(f"row: {i}, amt: {amt}")

# make copy of the selected columns
ep = ep.copy_col_name(['product', 'amount'])

# save the copy
ep.export('/some/dir/newfile.xlsx')
```
