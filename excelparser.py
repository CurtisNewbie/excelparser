from genericpath import exists
from typing import Callable
import pandas

'''
python3 -m pip install pandas openpyxl 

(openpyxl is optional, it's used for *.xlsx files)

Yongj.Zhuang
'''

isdebug = False
def debug(callback: Callable[[], str]):
    if isdebug: print("[debug] " + callback())


class ExcelParser():

    def __init__(self, inputf=None) -> None:
        self.inputf: str = inputf
        self.cols: list[str] = None
        self.rows: list[list[any]] = None
        self.cols_idx: dict[str][int] = None

    def append_row(self, row: list[any]):
        diff = len(self.cols) - len(row)
        if diff > 0:
            for i in range(diff):
                row.append("")
        self.rows.append(row)

    def append_empty_row(self):
        self.append_row([])

    def row_count(self) -> int:
        return len(self.rows)

    def getcol(self, col_name: str, row_idx: int) -> str:
        colidx = self.lookupcol(col_name)
        if colidx == -1: return ""
        return self.rows[row_idx][colidx]
    
    def joincol(self, col_name: str, delimiter: str = ",", wrap_each = None) -> str:
        idx = self.lookupcol(col_name)
        if idx == -1: return "" 

        joined  = ""
        for i in range(len(self.rows)):
            v = self.rows[i][idx]
            if wrap_each != None: v = wrap_each(v)
            joined += v
            if i < len(self.rows) - 1:
                joined += delimiter
        return joined

    def sumgroup(self, col_name: str, summed_col: str) -> dict[str]:
        grouped: dict[str][float] = {}
        grpidx = self.lookupcol(col_name)
        if grpidx == -1: return grouped 
        sumidx = self.lookupcol(summed_col)
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

    def sumcol(self, col_name: str) -> float:
        colidx = self.lookupcol(col_name)
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

    def lookupcol(self, col_name: int) -> int:
        '''
        Find column index by name
        '''
        if col_name not in self.cols_idx:
            return -1

        return self.cols_idx[col_name]

    def cvt_col_name(self, col_name: int, converter):
        '''
        Convert column value 
        '''
        colidx = self.lookupcol(col_name)
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
            idx = self.lookupcol(copied_names[i])
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
        ip: str = self.inputf
        if ip == None:
            raise ValueError("Please specify input file")
        debug(lambda: f"input path: '{ip}'")

        # read and parse workbook
        if not exists(ip):
            raise ValueError(f"Input file '{ip}' not found")

        df: pandas.DataFrame 
        if ip.lower().endswith('.csv'):
            df = pandas.read_csv(ip, 0, dtype="string")  
        else: 
            df = pandas.read_excel(ip, 0, dtype="string") 
        # debug(lambda: f"opened workbook '{ip}', content: '{df}'")

        nrow = len(df)
        ncol = len(df.columns)
        debug(lambda: f"row count: {nrow}, col count: {ncol}")

        # columns
        cols = []
        cols_i = 0
        self.cols_idx: dict[str][int] = {}

        for i in range(ncol):
            h = df.columns[i]
            if not h:
                break

            sh = str(h).strip()
            cols.append(sh)
            self.cols_idx[sh] = cols_i
            cols_i += 1

        if isdebug:
            s = "Columns: "
            for i in range(len(cols)):
                s += f"[{i}] {cols[i]}"
                if i < len(cols) - 1:
                    s += ", "
            debug(lambda: s)

        # rows
        rows: list[list[str]] = []
        for i in range(nrow):
            r = []
            for j in range(ncol):
                v = df.iat[i, j]
                if pandas.isnull(v):
                    v = ""
                r.append(v.strip())

            rows.append(r)
            debug(lambda: f"row[{i}]: {r}")

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