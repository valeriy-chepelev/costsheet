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
from pprint import pp
from docxtpl import DocxTemplate


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
    persons = list()
    for index, row in table.iterrows():
        # detect person record by cyrillic in column 2
        if re.match('[А-ЯЁа-яё \\-]+', str(row[2])):
            name = ' '.join(str(row[2]).split() +
                            str(table.loc[index + 1, 1]).split())  # split & join to remove extra spaces
            # over-check name structure
            if not re.match('^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*(?:\\s[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*){1,2}$',
                            name):
                raise ValueError(f'Name not properly formatted: "{name}"')
            emp_num = str(row[1]).strip()
            if not re.match('^\\d+$', emp_num):
                raise ValueError(f'Employee "{name}" number not properly formatted: "{emp_num}"')
            spec = str(table.loc[index + 2, 1]).strip()  # speciality
            status = table.loc[index + 1, 9:].tolist()  # slice of status letters
            times = row[9:].to_list()  # slice of day-times
            times = [0 if math.isnan(t) else t for t in times]  # zero NaNs
            # check days count is the same
            if len(times) != len(status):
                raise ImportError(f'Timedata not match status for "{name}"')
            # build dictionary data structure
            persons.append({'name': name,
                            'num': emp_num,
                            'spec': spec,
                            'time_data': times,
                            'pres_data': status})
        elif not (type(row[2]) is float and math.isnan(row[2])):
            # if ceil 2 contain something but not a name
            logging.warning(msg := f'Cell [{index}, 2] contain strange data.')
            print(f'{Fore.RED}WARNING: {msg}{Style.RESET_ALL}')
    return persons


def import_projects_data(filename):
    return pd.read_excel(filename, index_col=0)


def user_match(user: str, name: str):
    return user.lower() in name.lower()


def export_sheet(context):
    doc = DocxTemplate('MyTemplate.docx')
    doc.render(context)
    doc.save('DocExample.docx')


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

    # load projects/persons

    # load and parse employee table

    # find and update persons names in persons/projects, reindex

    # recalculate projects costs according to employee total time
    # (only if total cost > total time)

    # build person data table: project at day = time (time=min(table day time, rest_project_time until rpt=0)
    # if total cost too less when time - can divide day_time

    # build context: [projects [persons]]

    # render and output

    context = {'header': 'Header',
               'persons': import_hr_table('TestTable.xlsx')}
    export_sheet(context)




if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'{Fore.RED}Execution error:{e}{Style.RESET_ALL}')
        logging.exception('Common error')
