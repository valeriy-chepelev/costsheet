from yandex_tracker_client import TrackerClient
import configparser
import logging
import pandas as pd

LOG_LEVEL = logging.INFO


def main():

    # start logging

    logging.basicConfig(filename='costsheet.log',
                        filemode='a',
                        format='%(asctime)s %(name)s %(levelname)s %(message)s',
                        datefmt='%d/%m/%y %H:%M:%S',
                        level=LOG_LEVEL)
    logging.info('Costtrack started.')

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

    """

    # reading persons data

    persons = pd.read_excel('persons.xlsx')
    """

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
        input('Press any key to close...')