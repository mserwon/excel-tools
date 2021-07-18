"""
Analyze Gateway Export spreadsheets
Inspired by https://pbpython.com/excel-diff-pandas-update.html
For the documentation, download this file and type:
python compare.py --help
"""

import argparse

import pandas as pd
import numpy as np

def parse_excel(
        path1, sheet1, path2, sheet2, out_path, index_col_name, **kwargs
):
    p1 = path1.split('.')
    p2 = path2.split('.')
    if (p1[1] == 'csv'):
        old_df = pd.read_csv(path1, **kwargs)
    else:
        old_df = pd.read_excel(path1, sheet_name=sheet1, **kwargs)

    pd.loc[pd[]
    print(f"Differences saved in {out_path}")

def build_parser():
    cfg = argparse.ArgumentParser(
        description="Compares two excel files and outputs the differences "
                    "in another excel file."
    )
    cfg.add_argument("path1", help="Fist excel file")
    cfg.add_argument("sheet1", help="Name 1st excel sheet to compare.")
    cfg.add_argument("path2", help="Second excel file")
    cfg.add_argument("sheet2", help="Name 2nd excel sheet to compare.")
    cfg.add_argument(
        "key_column",
        help="Name of the column(s) with unique row identifier. It has to be "
             "the actual text of the first row, not the excel notation."
             "Use multiple times to create a composite index.",
        nargs="+",
    )
    cfg.add_argument("-o", "--output-path", default="compared.xlsx",
                     help="Path of the comparison results")
    cfg.add_argument("-m", "--merge-path", default="merged.xlsx",
                     help="Path of the merged results")
    cfg.add_argument("--skiprows", help='number of rows to skip', type=int,
                     action='append', default=None)
    return cfg


def main():
    cfg = build_parser()
    opt = cfg.parse_args()
    parse_excel(opt.path1, opt.sheet1, opt.path2, opt.sheet2, opt.output_path,
                  opt.key_column, skiprows=opt.skiprows)
    #merge_excel(opt.path1, opt.sheet1, opt.path2, opt.sheet2, opt.merge_path,
                  #opt.key_column, skiprows=opt.skiprows)

if __name__ == '__main__':
    main()