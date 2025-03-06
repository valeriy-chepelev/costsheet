from yandex_tracker_client import TrackerClient
import configparser
import logging

LOG_LEVEL = logging.INFO


def main():
    # start logging
    logging.basicConfig(filename='costsheet.log',
                        filemode='a',
                        format='%(asctime)s %(name)s %(levelname)s %(message)s',
                        datefmt='%d/%m/%y %H:%M:%S',
                        level=LOG_LEVEL)
    logging.info('Costtrack started.')
    # tracker connection
    config = configparser.ConfigParser()
    config.read('connect.ini')
    assert 'token' in config['DEFAULT']
    assert 'org' in config['DEFAULT']
    creds = config['DEFAULT']
    client = TrackerClient(creds['token'], creds['org'])
    if client.myself is None:
        raise Exception('Unable to connect Yandex Tracker.')
    # reading projects and persons
    config.read('costtrack.ini')
    persons = config['PERSONS']
    projects = config['PROJECTS']
    print(dict(persons))
    pass


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('Execution error:', e)
        logging.exception('Common error')
        input('Press any key to close...')