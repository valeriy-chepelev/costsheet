import logging
import pandas as pd
import datetime as dt
import argparse
import os
from colorama import init as colorama_init
from colorama import Fore
from colorama import Style
import re
import math
from docxtpl import DocxTemplate


def define_parser():
    """ Return CLI arguments parser
    """
    parser = argparse.ArgumentParser(description='Costsheet|Costsheet v.1.0 - Costs formatter by VCh.')
    parser.add_argument('table_filename', nargs='?',
                        default='T13_default_table_name_magic_value',
                        help='monthly employees report xlsx filename; default - "T13-yy-mm.xlsx"')
    parser.add_argument('-t', '--template', metavar='TEMPLATE', default='t-13-template v5.docx',
                        help='template docx filename (default "t-13-template v5.docx"')
    parser.add_argument('-d', '--date', metavar='REPORT_DATE',
                        type=lambda s: dt.datetime.strptime(s, '%y-%m'),
                        help='report period in "y-m" format (like "25-1" for january 2025); '
                             'default - previous month until 14th, current month since 15th')
    parser.add_argument('--debug', default=False, action='store_true',
                        help='logging in debug mode')
    return parser


def import_hr_table(filename):
    in_table = pd.read_excel(filename, header=None, index_col=None,
                             skiprows=9, skipfooter=5)
    persons = list()
    for index, row in in_table.iterrows():
        # detect person record by cyrillic in column 2
        if re.match('[А-ЯЁа-яё \\-]+', str(row[2])):
            name = ' '.join(str(row[2]).split() +
                            str(in_table.iloc[int(index) + 1, 1]).split())  # split & join to remove extra spaces
            # over-check name structure
            if not re.match('^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*(?:\\s[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)*){1,2}$',
                            name):
                raise ValueError(f'Name not properly formatted: "{name}"')
            emp_num = str(row[1]).strip()
            if not re.match('^\\d+$', emp_num):
                raise ValueError(f'Employee "{name}" number not properly formatted: "{emp_num}"')
            spec = str(in_table.iloc[int(index) + 2, 1]).strip()  # speciality
            status = in_table.iloc[int(index) + 1, 9:].tolist()  # slice of status letters
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


def export_sheet(context, template, filename='DocExample.docx'):
    print(f'{Fore.GREEN}Rendering...{Style.RESET_ALL}')
    doc = DocxTemplate(template)
    doc.render(context)
    doc.save(filename)
    print(f'{Fore.GREEN}Complete and stored to "{filename}".{Style.RESET_ALL}')


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
    today = dt.datetime.now(dt.timezone.utc)
    rep_period = today + dt.timedelta(days=-14) if args.date is None else args.date
    days_count = ((rep_period.replace(day=28) + dt.timedelta(days=4)).replace(day=1) + dt.timedelta(days=-1)).day
    print(f'Building costs report for {Fore.GREEN}{rep_period.strftime("%B %Y")}{Style.RESET_ALL}')

    # load projects/persons and boss
    pp_filename = f'costs-{rep_period.strftime("%y-%m")}.xlsx'
    if not os.path.isfile(pp_filename):
        raise ValueError(f'{pp_filename} not found!')
    pp_costs = pd.read_excel(pp_filename, sheet_name='costs', index_col=0)
    boss = pd.read_excel(pp_filename, sheet_name='boss', header=None, index_col=None, usecols=[1])
    print(boss.loc[0, 1], boss.loc[1, 1])

    # load and parse employee table
    if args.table_filename == 'T13_default_table_name_magic_value':
        emp_filename = f'T13-{rep_period.strftime("%y-%m")}.xlsx'
    else:
        emp_filename = args.table_filename
    emp_table = pd.DataFrame(import_hr_table(emp_filename))
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
    pp_costs.set_index('fullname', inplace=True)  # full names added and reindex
    pp_costs['summary'] = pp_costs.sum(axis='columns')  # get summary column
    pp_costs = pp_costs.loc[(pp_costs.sum(axis=1) != 0), (pp_costs.sum(axis=0) != 0)]  # drop zero rows and cols
    pp_costs.sort_index()  # sort by person name

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
            while not project_cost < 1:  # add project cost to days, day by day
                spent = int(min(project_cost, job))  # how much we can spend to a day
                if spent:
                    pers['projects'][project_name].update({f'h{date}': spent,
                                                           f'pres{date}': presence})  # add to project
                    project_cost -= spent  # decrease project cost
                    job -= spent  # decrease day job
                if job < 1 and not project_cost < 1:  # if day job over, but need continue - take next day
                    date += 1
                    job = emp_data[f'h{date}']
                    presence = emp_data[f'pres{date}']
            # calc half-totals and totals
            part1 = [val for key, val in pers['projects'][project_name].items()
                     if re.match('^h(?:1[0-5]|[1-9])$', key) and type(val) is int]
            part2 = [val for key, val in pers['projects'][project_name].items()
                     if re.match('^h(1[6-9]|2\\d|3[0-1])$', key) and type(val) is int]
            pers['projects'][project_name].update({
                'hp1': sum(part1), 'dp1': len(part1),
                'hp2': sum(part2), 'dp2': len(part2),
                'sh': sum(part1) + sum(part2), 'sd': len(part1) + len(part2)})
        pers_list.append(pers)

    # fill empty days with default employee presence and zero times

    for pers_data in pers_list:
        for project in pers_data['projects']:
            for date in range(1, 32):  # iterate max number of days: 1 to 31
                if f'h{date}' not in pers_data['projects'][project]:  # is date empty
                    try:  # if days exceeds month - we get KeyError from emp_table.loc
                        if date > days_count:
                            raise KeyError('outdated')
                        pers_data['projects'][project].update(
                            {f'h{date}': ' ',
                             f'pres{date}': 'Н' if emp_table.loc[pers_data['name'], f'h{date}'] > 0
                             else emp_table.loc[pers_data['name'], f'pres{date}']})
                    except KeyError:
                        pers_data['projects'][project].update(
                            {f'h{date}': 'X', f'pres{date}': 'X'})

    # build context: [projects [persons]]

    common_dict = {'rep_date': today.strftime("%d.%m.%Y"),
                   'rep_period': rep_period.strftime("%m месяц %Y год"),
                   'hod_spec': boss.loc[0, 1],
                   'hod_name': boss.loc[1, 1]}
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
    report_filename = ''.join(s for s in os.getlogin() if s.isalnum()) + f'-t13-{rep_period.strftime("%y-%m")}.docx'
    export_sheet(context, args.template, report_filename)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'{Fore.RED}Execution error:{e}{Style.RESET_ALL}')
        logging.exception('Common error')
    input('Press "Enter" to leave.')
