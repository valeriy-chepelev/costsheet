from yandex_tracker_client import TrackerClient
import configparser
import logging
import pandas as pd
from alive_progress import alive_bar
import datetime as dt
import argparse


def define_parser():
    """ Return CLI arguments parser
    """
    parser = argparse.ArgumentParser(description='Costsheet|Costtrack v.1.0 - Yandex Tracker costs crawler by VCh.',
                                     epilog='Tracker connection settings in "connect.ini".')

    parser.add_argument('-d', '--date', metavar='REPORT_DATE',
                        type=lambda s: dt.datetime.strptime(s, '%m-%y'),
                        help='specify report date in "m-y" format (like "1-25" for january 2025)')
    parser.add_argument('--debug', default=False, action='store_true',
                        help='logging in debug mode (include tracker and issues info)')
    return parser


def spend(issue, person, date):
    # TODO: calculate
    return 8


def login_match(login, user) -> bool:
    return login.lower() in user.display.lower()


def main():
    args = define_parser().parse_args()  # get CLI arguments

    # start logging

    logging.basicConfig(filename='costsheet.log',
                        filemode='a',
                        format='%(asctime)s %(name)s %(levelname)s %(message)s',
                        datefmt='%d/%m/%y %H:%M:%S',
                        level=logging.INFO if args.debug else logging.ERROR)
    logging.info('Costtrack started.')

    # Configure pandas to full output

    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    # define dates
    # in not defined in argument - two last week set to next month, otherwise - to current month

    date = dt.datetime.now(dt.timezone.utc) + dt.timedelta(days=14) if args.date is None else args.date

    start_date = date.replace(day=1)
    final_date = (date.replace(day=28) + dt.timedelta(days=4)).replace(day=1) + dt.timedelta(days=-1)
    print(f'Gathering costs report for {date.strftime("%B %Y")}')

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

    projects = pd.read_excel('ScanData.xlsx',
                             sheet_name='Projects',
                             header=None, index_col=None,
                             usecols=[0, 1], skiprows=1,
                             names=['name', 'request'])
    print('Projects list:\n', projects)

    # reading persons data

    persons = pd.read_excel('ScanData.xlsx',
                            sheet_name='Persons',
                            header=None, index_col=None,
                            usecols=[0, 1, 2], skiprows=1,
                            names=['name', 'login', 'lazy'])

    # lookup persons in Tracker

    persons['accounts'] = ''
    with alive_bar(len(persons), title='Persons lookup', theme='classic') as bar:
        for index_pers, person in persons.iterrows():
            acc = [user.display for user in client.users
                   if login_match(person['login'], user)]
            persons.at[index_pers, 'accounts'] = ';'.join(acc) if len(acc) else 'WARNING: no accounts!'
            bar()
    print('Persons list:\n', persons)

    # acquiring data from Tracker

    for index_prj, project in projects.iterrows():
        issues = client.issues.find(query=project['request'])
        print(f"Project '{project['name']}' contain {len(issues)} issue(s).")
        """
        for index_pers, person in persons.iterrows():
            s = 0
            for issue in issues:
                 # TODO: summ spends


    # store the report

    pass
    
    """


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('Execution error:', e)
        logging.exception('Common error')
