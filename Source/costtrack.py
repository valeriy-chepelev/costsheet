from yandex_tracker_client import TrackerClient
import configparser
import logging
import pandas as pd
from alive_progress import alive_bar
import datetime as dt
import argparse
import os
from data_access import issue_times, iso_hrs
from colorama import init as colorama_init
from colorama import Fore
from colorama import Style


def define_parser():
    """ Return CLI arguments parser
    """
    parser = argparse.ArgumentParser(description='Costsheet|Costtrack v.1.0 - Yandex Tracker costs crawler by VCh.',
                                     epilog='Tracker connection settings in "connect.ini".')
    parser.add_argument('filename', nargs='?', default='ScanData.xlsx',
                        help='input excel projects and persons config; default - "ScanData.xlsx"')
    parser.add_argument('-d', '--date', metavar='REPORT_DATE',
                        type=lambda s: dt.datetime.strptime(s, '%y-%m'),
                        help='specify report date in "y-m" format (like "25-1" for january 2025); '
                             'default - previous month until 14th, current month since 15th')
    parser.add_argument('-n', '--noname', dest='add_name', action='store_false',
                        help='suppress user name in report filename')
    parser.add_argument('--debug', default=False, action='store_true',
                        help='logging in debug mode (include tracker and issues info)')
    parser.set_defaults(add_name=True)
    return parser


def spend(issue, person, start, final):
    try:
        lc = int(person['move_cost'])
    except ValueError:
        lc = 0
    sp = sum(((iso_hrs(x['value']) - iso_hrs(x['from'])) if x['kind'] == 'spent' else lc)
             for x in issue_times(issue)
             if login_match(person['login'], x['by']) and
             x['kind'] in ['spent', 'status'] and
             start.date() <= x['date'].date() <= final.date()
             )
    return sp


def login_match(login, user) -> bool:
    try:
        u = user.display
    except AttributeError:
        u = user
    return login.lower() in u.lower()


def users_jaccard(users):
    """
    Test difference between usernames
    @param users: list of usernames
    @return: min jaccard factor (float <= 1) of username relative to all the symbols
    """
    union = set(''.join(users))  # all the characters
    return min([len(set(user) & union) / len(set(user) | union) for user in users], default=1.0)


def main():
    # init

    args = define_parser().parse_args()  # get CLI arguments
    colorama_init()
    logging.basicConfig(filename='costsheet.log',
                        filemode='a',
                        format='%(asctime)s %(name)s %(levelname)s %(message)s',
                        datefmt='%d/%m/%y %H:%M:%S',
                        level=logging.INFO if args.debug else logging.WARNING)
    logging.info('Costtrack started.')

    # Configure pandas to full output

    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    # check input file present

    if not os.path.isfile(args.filename):
        raise ValueError(f'{args.filename} is not a file!')

    # define dates
    # in not defined in argument - two first week set to previous month, otherwise - to current month

    date = dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=-14) if args.date is None else args.date
    start_date = date.replace(day=1)
    final_date = (date.replace(day=28) + dt.timedelta(days=4)).replace(day=1) + dt.timedelta(days=-1)
    print(f'Gathering costs report for {Fore.GREEN}{date.strftime("%B %Y")}{Style.RESET_ALL}')

    # configure and establish Tracker connection

    config = configparser.ConfigParser()
    config.read('connect.ini')
    assert 'token' in config['DEFAULT']
    assert 'org' in config['DEFAULT']
    creds = config['DEFAULT']
    client = TrackerClient(creds['token'], creds['org'])
    if client.myself is None:
        raise Exception('Unable to connect Yandex Tracker.')

    # reading projects data

    projects = pd.read_excel(args.filename,
                             sheet_name='Projects',
                             header=None, index_col=None,
                             usecols=[0, 1], skiprows=1,
                             names=['name', 'request'])
    print()
    projects['size'] = 0
    with alive_bar(len(projects), title='Projects', theme='classic') as bar:
        for index_prj, project in projects.iterrows():
            projects.at[index_prj, 'size'] = len(list(client.issues.find(query=project['request'])))
            bar()
    print(projects)

    # reading persons data

    persons = pd.read_excel(args.filename,
                            sheet_name='Persons',
                            header=None, index_col=None,
                            usecols=[0, 1, 2], skiprows=1,
                            names=['name', 'login', 'move_cost'])
    print()
    persons['accounts'] = ''
    with alive_bar(len(persons), title='Persons', theme='classic') as bar:
        for index_pers, person in persons.iterrows():
            users_list = [user.display for user in client.users if login_match(person['login'], user)]
            jf = users_jaccard([a.split('@')[0].lower() for a in users_list])
            warn = ''
            if jf < 0.8:
                warn = f'{Fore.RED}WARNING: too wide selector: {Style.RESET_ALL}'
                logging.warning(f'too wide selector as {person["login"]}: {";".join(users_list)}')
            if len(users_list):
                persons.at[index_pers, 'accounts'] = warn + ';'.join(users_list)
            else:
                logging.warning(f'no accounts for {person["login"]}')
                persons.at[index_pers, 'accounts'] = f'{Fore.RED}WARNING: no accounts!{Style.RESET_ALL}'
            bar()
    print(persons)

    # acquiring data from Tracker

    print()
    report = pd.DataFrame(0,
                          index=persons['name'].values.tolist(),
                          columns=projects['name'].values.tolist())
    with alive_bar(int(len(persons) * sum(projects['size'].values)),
                   title='Costs', theme='classic') as bar:
        for _, project in projects.iterrows():
            issues = client.issues.find(query=project['request'])
            for _, person in persons.iterrows():
                s = 0
                for issue in issues:
                    if (dt.datetime.strptime(issue.updatedAt, '%Y-%m-%dT%H:%M:%S.%f%z').date()
                        >= start_date.date()) and (
                            dt.datetime.strptime(issue.createdAt, '%Y-%m-%dT%H:%M:%S.%f%z').date()
                            <= final_date.date()):
                        s += spend(issue, person, start_date, final_date)
                    bar()
                report.at[person['name'], project['name']] = s
    print(report)

    # store the report

    user_name = ''.join(s for s in os.getlogin() if s.isalnum()) + '-'
    report_name = f'{user_name if args.add_name else ""}costs-{date.strftime("%y-%m")}.xlsx'
    with pd.ExcelWriter(report_name) as writer:
        report.to_excel(writer, sheet_name='costs')
        # use writer to save multiple dataframes as sheets in one file
    print(f'\n{Fore.GREEN}Report successfully stored to {Fore.CYAN}{report_name}{Style.RESET_ALL}')


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f'{Fore.RED}Execution error:{e}{Style.RESET_ALL}')
        logging.exception('Common error')
