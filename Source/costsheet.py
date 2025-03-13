import logging
import pandas as pd
from alive_progress import alive_bar
import datetime as dt
import argparse
import os
from colorama import init as colorama_init
from colorama import Fore
from colorama import Style
import re
import math


def define_parser():
    """ Return CLI arguments parser
    """
    parser = argparse.ArgumentParser(description='Costsheet|Costsheet v.1.0 - Costs formatter by VCh.')
    parser.add_argument('--debug', default=False, action='store_true',
                        help='logging in debug mode (include tracker and issues info)')
    return parser


def import_hr_table(filename):
    table = pd.read_excel(filename, header=None, index_col=None,
                          skiprows=9, skipfooter=5)
    for index, row in table.iterrows():
        # detect person record by cyrillic in column 2
        if re.match('[А-ЯЁа-яё \\-]+', str(row[2])):
            name = ' '.join(str(row[2]).split() +
                            str(table.loc[index + 1, 1]).split())  # split & join to remove extra spaces
            # over-check name structure
            if not re.match('^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*(?:\\s[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*){1,2}$',
                            name):
                raise ValueError(f'Name not properly formatted: "{name}"')
            spec = str(table.loc[index + 2, 1]).strip()
            status = table.loc[index + 1, 9:].tolist()
            times = row[9:].to_list()
            times = [0 if math.isnan(t) else t for t in times]  # Zero Nans
            if len(times) != len(status):
                raise ImportError(f'Timedata not match status for "{name}"')
            print(name, spec)
            print(times)
            print(status)


def import_projects_data():
    pass


def export_sheet():
    pass


def main():
    # init

    args = define_parser().parse_args()  # get CLI arguments
    colorama_init()
    logging.basicConfig(filename='costsheet.log',
                        filemode='a',
                        format='%(asctime)s %(name)s %(levelname)s %(message)s',
                        datefmt='%d/%m/%y %H:%M:%S',
                        level=logging.INFO if args.debug else logging.WARNING)
    logging.info('Costsheet started.')

    # Configure pandas to full output

    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    import_hr_table('TestTable.xlsx')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'{Fore.RED}Execution error:{e}{Style.RESET_ALL}')
        logging.exception('Common error')
