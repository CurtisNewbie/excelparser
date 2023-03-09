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
        self.inputf = inputf
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
        if colidx == -1:
            return ""
        return self.rows[row_idx][colidx]

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

    def parse(self):
        '''
        Parse excel
        '''
        ip = self.inputf
        if ip == None:
            raise ValueError("Please specify input file")
        debug(lambda: f"input path: '{ip}'")

        # read and parse workbook
        if not exists(ip):
            raise ValueError(f"Input file '{ip}' not found")

        df: pandas.DataFrame = pandas.read_excel(ip, 0, dtype="string")
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

