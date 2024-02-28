from genericpath import exists
import pandas

class RowIter():
    def __init__(self, p: "ExcelParser"):
        self.i = -1
        self.p = p

    def __next__(self):
        self.i += 1
        if self.i < len(self.p.rows):
            return Row(self.i, self.p)
        raise StopIteration

class Row:
    def __init__(self, i: int, p: "ExcelParser"):
        self.i = i
        self.p = p

    def get_col(self, col_name: str):
        return self.p.get_col(col_name, self.i)

    def get_col_at(self, idx: int):
        return self.p.get_col_at(idx, self.i)

class ExcelParser():

    def __init__(self, inputf=None) -> None:
        self.inputf: str = inputf
        self.cols: list[str] = []
        self.rows: list[list[any]] = []
        self.cols_idx: dict[str][int] = {}

    def __iter__(self):
        return RowIter(self)

    def append_row(self, row: list[any]):
        '''
        Append row

        If the len of row is less than len of columns, the remaining is filled with empty string
        '''
        diff = len(self.cols) - len(row)
        if diff > 0:
            for i in range(diff): row.append("")
        self.rows.append(row)

    def flatten_col(self, col_name: str) -> list[str]:
        '''
        Get all values of the specified column as a list
        '''
        colidx = self.lookup_col(col_name)
        if colidx == -1: return []
        l = []
        for r in self.rows: l.append(r[colidx])
        return l

    def flatten_all(self) -> list[str]:
        '''
        Get all values of cols and rows as a list
        '''
        l = []
        for c in self.cols: l.append(c)
        for r in self.rows:
            for v in r: l.append(v)
        return l

    def flatten_rows(self) -> list[str]:
        '''
        Get all values of rows as a list
        '''
        l = []
        for r in self.rows:
            for v in r: l.append(v)
        return l

    def append_empty_row(self):
        '''
        Append empty row
        '''
        self.append_row([])

    def row_count(self) -> int:
        '''
        Return len(rows)
        '''
        return len(self.rows)

    def get_col(self, col_name: str, row_idx: int) -> str:
        '''
        Get column value for the row
        '''
        colidx = self.lookup_col(col_name)
        if colidx == -1: return ""
        return self.rows[row_idx][colidx]

    def fuzz_find_col(self, search: str) -> str:
        for c in self.cols:
            if search in c: return c

    def get_col_at(self, idx: int, row_idx: int) -> str:
        '''
        Get column value for the row at specific column
        '''
        return self.rows[row_idx][idx]

    def collect_col(self, col_name: str) -> set[str]:
        '''
        Collect all values of the column to a set
        '''
        return set(self.flatten_col(col_name))

    def filter_rows(self, filter):
        '''
        filter rows

        filter is a function that take a rows (list[str]) and return boolean, a row is filtered if it returns False
        '''
        newrow: list[list[any]] = []
        for r in self.rows:
            if not filter(r): newrow.append(r)
        self.rows = newrow

    def append_cols(self, col_names: list[str]):
        '''
        Append new columns at the end
        '''
        for c in col_names:
            self.append_col(c)

    def append_col(self, col_name: str):
        '''
        Append a new column at the end
        '''
        self.append_col_val(col_name, lambda : "")

    def append_col_val(self, col_name: str, value_supplier):
        '''
        Append a new column at the end

        value_supplier is a function that returns random value
        '''
        self.cols.append(col_name)
        self.cols_idx[col_name] = len(self.cols) - 1
        for r in self.rows: r.append(value_supplier())

    def quote_join_col(self, col_name: str, delimiter: str = ",") -> str:
        '''
        Join column values with the delimiter, and wrap each value with single quote
        '''
        return self.join_col(col_name, delimiter, quote)

    def join_col(self, col_name: str, delimiter: str = ",", wrap_each = None) -> str:
        '''
        Join column values with the delimiter, and wrap each value if necessary
        '''
        idx = self.lookup_col(col_name)
        if idx == -1: return ""

        joined  = ""
        for i in range(len(self.rows)):
            v = self.rows[i][idx]
            if wrap_each != None: v = wrap_each(v)
            joined += v
            if i < len(self.rows) - 1:
                joined += delimiter
        return joined

    def group_distinct(self, col_name: str) -> dict[str]:
        '''
        Group rows by distinct column value

        key: column value
        value: row index
        '''
        grouped: dict[str][int] = {}
        idx = self.lookup_col(col_name)
        if idx == -1:
            raise ValueError(f"Column {col_name} is not found")

        for i in range(len(self.rows)):
            k = self.rows[i][idx]
            grouped[k] = i

        return grouped


    def group_and_sum(self, col_name: str, summed_col: str) -> dict[str]:
        '''
        Group and sum values of the specified column
        '''
        grouped: dict[str][float] = {}
        grpidx = self.lookup_col(col_name)
        if grpidx == -1: return grouped
        sumidx = self.lookup_col(summed_col)
        if sumidx == -1: return grouped

        for i in range(len(self.rows)):
            k = self.rows[i][grpidx]
            if k not in grouped:
                grouped[k] = 0

            v = self.rows[i][sumidx]
            if v == "":
                v = 0
            else:
                v = float(v)
            grouped[k] += v

        return grouped

    def sum_col(self, col_name: str) -> float:
        '''
        Sum values of column
        '''
        colidx = self.lookup_col(col_name)
        if colidx == -1:
            return 0

        sum = 0
        for i in range(len(self.rows)):
            v = self.rows[i][colidx]
            if v == "":
                v = 0
            else:
                v = float(v)
            sum += v
        return sum

    def row_get_col(self, col_name: str, row: list[str]) -> str:
        '''
        Find column value in row by name
        '''
        i = self.lookup_col(col_name)
        if i < 0: return ""
        return row[i]

    def lookup_col(self, col_name: str) -> int:
        '''
        Find column index by name
        '''
        if col_name not in self.cols_idx: return -1
        return self.cols_idx[col_name]

    def cvt_col(self, col_name: str, converter):
        '''
        Convert column value
        '''
        colidx = self.lookup_col(col_name)
        if colidx == -1:
            raise ValueError(
                f"Unable to find '{col_name}', available columns are: {self.cols}")
        self.cvt_col_at(colidx, converter)

    def cvt_col_at(self, col_idx: int, converter):
        '''
        Convert column value
        '''
        for i in range(len(self.rows)):
            self.rows[i][col_idx] = converter(self.rows[i][col_idx])

    def copy_col_name(self, copied_names: list[int]) -> "ExcelParser":
        '''
        Copy columns
        '''
        colnames = []
        idxls: list[int] = []

        for i in range(len(copied_names)):
            idx = self.lookup_col(copied_names[i])
            if idx > -1:
                colnames.append(copied_names[i])
                idxls.append(idx)

        ep = ExcelParser()
        if len(idxls) > 0:
            ep.rows = []
            ep.cols = colnames

            for i in range(len(self.rows)):
                r: list = self.rows[i]
                cprow = []
                for i in range(len(idxls)):
                    cprow.append(r[idxls[i]])
                ep.rows.append(cprow)

        return ep

    def export(self, outf):
        '''
        Export to file
        '''
        df = pandas.DataFrame(self.rows, columns=self.cols)
        df.to_excel(outf, index=False)

    def parse(self) -> "ExcelParser":
        '''
        Parse file
        '''
        ip: str = self.inputf
        if ip == None: raise ValueError("Please specify input file")

        # read and parse workbook
        if not exists(ip): raise ValueError(f"Input file '{ip}' not found")

        df: pandas.DataFrame
        if ip.lower().endswith('.csv'): df = pandas.read_csv(ip, 0, dtype="string")
        else: df = pandas.read_excel(ip, 0, dtype="string")

        nrow = len(df)
        ncol = len(df.columns)

        # columns
        cols = []
        cols_i = 0
        self.cols_idx: dict[str][int] = {}

        for i in range(ncol):
            h = df.columns[i]
            if not h: break

            sh = str(h).strip()
            cols.append(sh)
            self.cols_idx[sh] = cols_i
            cols_i += 1

        # rows
        rows: list[list[str]] = []
        for i in range(nrow):
            r = []
            for j in range(ncol):
                v = df.iat[i, j]
                if pandas.isnull(v): v = ""
                r.append(v.strip())

            rows.append(r)

        self.cols = cols
        self.rows = rows
        return self

    def __str__(self) -> str:
        s = f"_File: '{self.inputf}'"
        s += "\n_Columns: " + str(self.cols)
        s += "\n_Rows: "
        for i in range(len(self.rows)):
            s += "\n" + str(self.rows[i])
        return s

    # >>>>>>>>>>>>>>>>>>>> deprecated stuff >>>>>>>>>>>>>>>>>

    def getcol(self, col_name: str, row_idx: int) -> str:
        '''
        deprecated, use get_col instead
        '''
        return self.get_col(col_name, row_idx)

    def appendcols(self, col_names: list[str]):
        '''
        deprecated, use append_cols instead
        '''
        self.append_cols(col_names)

    def appendcol(self, col_name: str):
        '''
        deprecated, use appendcol instead
        '''
        self.append_col(col_name)

    def appendcolval(self, col_name: str, value_supplier):
        '''
        deprecated, use append_col_val instead
        '''
        self.append_col_val(col_name, value_supplier)

    def cvt_col_name(self, col_name: str, converter):
        '''
        deprecated, use cvt_col instead
        '''
        return self.cvt_col(col_name, converter)

    def row_lookup_col(self, col_name: str, row: list[str]) -> str:
        '''
        deprecated, use row_get_col instead
        '''
        return self.row_get_col(col_name, row)

    # <<<<<<<<<<<<<<<<<<<< deprecated stuff <<<<<<<<<<<<<<<<<

def quote_each(l: list[str]) -> list[str]:
    return [quote(v) for v in l]

def quote(s: str) -> str:
    return "'" + s + "'"

def parse(f: str) -> ExcelParser:
    '''
    Same as ExcelParser(f).parse()
    '''
    return ExcelParser(f).parse()

def empty_parser(*col: str) -> ExcelParser:
    '''
    Create empty ExcelParser with provided columns
    '''
    p = ExcelParser()
    p.append_cols(col)
    return p
