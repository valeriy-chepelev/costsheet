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
import math


def define_parser():
    """ Return CLI arguments parser
    """
    parser = argparse.ArgumentParser(description='Costsheet|Costsheet v.1.0 - Costs formatter by VCh.')
    parser.add_argument('-t', '--template', metavar='TEMPLATE', default='t-13-template v4.docx',
                        help='template docx filename (default "tt-13-template v4.docx"')
    parser.add_argument('-e', '--employees', metavar='EMPLOYEES', default='TestTable.xlsx',
                        help='employees monthly report xlsx filename (default "TestTable.xlsx"')
    parser.add_argument('-d', '--date', metavar='REPORT_DATE',
                        type=lambda s: dt.datetime.strptime(s, '%y-%m'),
                        help='specify report period in "y-m" format (like "25-1" for january 2025); '
                             'default - previous month until 14th, current month since 15th')
    parser.add_argument('--debug', default=False, action='store_true',
                        help='logging in debug mode (include tracker and issues info)')
    return parser


def import_hr_table(filename):
    in_table = pd.read_excel(filename, header=None, index_col=None,
                             skiprows=9, skipfooter=5)
    persons = list()
    for index, row in in_table.iterrows():
        # detect person record by cyrillic in column 2
        if re.match('[А-ЯЁа-яё \\-]+', str(row[2])):
            name = ' '.join(str(row[2]).split() +
                            str(in_table.loc[index + 1, 1]).split())  # split & join to remove extra spaces
            # over-check name structure
            if not re.match('^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*(?:\\s[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*){1,2}$',
                            name):
                raise ValueError(f'Name not properly formatted: "{name}"')
            emp_num = str(row[1]).strip()
            if not re.match('^\\d+$', emp_num):
                raise ValueError(f'Employee "{name}" number not properly formatted: "{emp_num}"')
            spec = str(in_table.loc[index + 2, 1]).strip()  # speciality
            status = in_table.loc[index + 1, 9:].tolist()  # slice of status letters
            times = row[9:].to_list()  # slice of day-times
            times = [0 if math.isnan(t) else t for t in times]  # zero NaNs
            # check days count is the same
            if len(times) != len(status):
                raise ImportError(f'Timedata not match status for "{name}"')
            # build dictionary data structure
            person = {'name': name, 'num': emp_num, 'spec': spec, 'total': sum(times)}
            person.update({f'h{i}': t for i, t in enumerate(times, 1)})
            person.update({f'pres{i}': t for i, t in enumerate(status, 1)})
            persons.append(person)
        elif not (type(row[2]) is float and math.isnan(row[2])):
            # if ceil 2 contain something but not a name
            logging.warning(msg := f'Cell [{index}, 2] contain strange data.')
            print(f'{Fore.RED}WARNING: {msg}{Style.RESET_ALL}')
    return persons


def export_sheet(context, template):
    print('Rendering...')
    doc = DocxTemplate(template)
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

    # define dates
    # in not defined in argument - two first week set to previous month, otherwise - to current month
    date = dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=-14) if args.date is None else args.date
    print(f'Building costs report for {Fore.GREEN}{date.strftime("%B %Y")}{Style.RESET_ALL}')

    # load projects/persons and boss
    pp_filename = f'costs-{date.strftime("%y-%m")}.xlsx'
    if not os.path.isfile(pp_filename):
        raise ValueError(f'{pp_filename} not found!')
    pp_costs = pd.read_excel(pp_filename, sheet_name='costs', index_col=0)
    boss = pd.read_excel(pp_filename, sheet_name='boss', header=None, index_col=None, usecols=[1])
    print(boss.loc[0, 1], boss.loc[1, 1])

    # load and parse employee table
    emp_table = pd.DataFrame(import_hr_table(args.employees))
    emp_table.set_index('name', inplace=True)

    # find and update persons names in persons/projects
    # reindex, clear zero persons, sort by name
    emp_names = emp_table.index.tolist()  # list of full names
    full_names = list()
    for name in pp_costs.index.tolist():
        new_name = [n for n in emp_names if name.lower() in n.lower()]  # get all matching full names
        if len(new_name) > 1:
            print(f'{Fore.RED}WARNING: too wide selector for {name}: {"; ".join(new_name)}{Style.RESET_ALL}')
            logging.warning(f'too wide selector as {name}: {";".join(new_name)}')
        if len(new_name) == 0:
            raise ValueError(f'no employee data for {name}')
        full_names.append(new_name[0])
    pp_costs['fullname'] = full_names
    pp_costs.set_index('fullname', inplace=True)  # full names added and reindexed
    pp_costs['summary'] = pp_costs.sum(axis='columns')  # get summary column
    pp_costs.drop(pp_costs[pp_costs.summary == 0].index, inplace=True)  # drop persons with zero summary
    # pp_costs.drop('summary', axis='columns', inplace=True)  # delete summary column ???
    pp_costs.sort_index()  # sort by name

    # calculate employee projects factor (total job time/ projects summary time)
    pp_costs = pp_costs.merge(emp_table[['total']], how='left', sort=False,
                              left_index=True, right_index=True).eval('factor=total/summary')
    # trim exceeded cost
    for person, data in pp_costs.iterrows():
        if data['factor'] < 1:
            print(f'{Fore.RED}WARNING: exceeded time for {person}: {data["summary"]} > {data["total"]},'
                  f' corrected. {Style.RESET_ALL}')
            logging.warning(f'exceeded time for {person}: {data["summary"]} > {data["total"]},'
                            f' corrected.')
            for project in list(pp_costs)[:-3]:
                pp_costs.at[person, project] = math.floor(data['factor'] * data[project])
    print('Costs data:')
    print(pp_costs)

    # build persons data structure: project at day = time (time=min(table day time, rest_project_time until rpt=0)

    prj_names = list(pp_costs)[:-3]
    pers_list = list()
    for person, costs in pp_costs.iterrows():  # iterate persons and projects costs
        emp_data = emp_table.loc[person]  # employee daily hours and presence
        days_count = (len(emp_data) - 3) // 2  # 3 is first position of hours data
        date = 1  # start at first day, select employee daily data
        job = emp_data[f'h{date}']
        presence = emp_data[f'pres{date}']
        # prepare person dictionary
        pers = {'name': person,
                'num': emp_data['num'],
                'spec': emp_data['spec'],
                'projects': dict()}
        # iterate projects those person participate
        for project_name, project_cost in [x for x in zip(prj_names, list(costs)[:-3]) if x[1] > 0]:
            pers['projects'].update({project_name: dict()})  # attach project to person
            while project_cost > 0:  # add project cost to days, day by day
                spent = min(project_cost, job)  # how much we can spent to a day
                pers['projects'][project_name].update({f'h{date}': spent,
                                                       f'pres{date}': presence})  # add to project
                project_cost -= spent  # decrease project cost
                job -= spent  # decrease day job
                if job < 1:  # if day job over - take next day
                    date += 1
                    job = emp_data[f'h{date}']
                    presence = emp_data[f'pres{date}']
        pers_list.append(pers)

    # fill empty days with default employee presence and zero times

    for pers_data in pers_list:
        for project in pers_data['projects']:
            for date in range(1, 32):  # iterate max number of days: 1 to 31
                if f'h{date}' not in pers_data['projects'][project]:  # is date empty
                    try:  # if days exceeds month - we get KeyError from emp_table.loc
                        pers_data['projects'][project].update(
                            {f'h{date}': 0,
                             f'pres{date}': emp_table.loc[pers_data['name'],
                             f'pres{date}']})
                    except KeyError:
                        pass

    # build context: [projects [persons]]

    common_dict = {'rep_date': 1,
                   'rep_period': 2,
                   'hod_spec': boss.loc[0, 1],
                   'hod_name': boss.loc[1, 1]}  # TODO: common context data
    context = {'projects': list()}
    for project in prj_names:
        ctx_pers_list = [p for p in pers_list if project in p['projects']]  # list filtered by project
        p_dict = {'project': project,
                  'emps': [{'ord': i,
                            'name': p['name'],
                            'num': p['num'],
                            'spec': p['spec']} | p['projects'][project]
                           for i, p in enumerate(ctx_pers_list, 1)]}
        p_dict.update(common_dict)
        context['projects'].append(p_dict)

    # render and output
    export_sheet(context, args.template)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'{Fore.RED}Execution error:{e}{Style.RESET_ALL}')
        logging.exception('Common error')
