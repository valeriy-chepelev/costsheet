from functools import lru_cache
import datetime as dt
from yandex_tracker_client.exceptions import Forbidden


def _iso_split(s, split):
    """ Splitter helper for converting ISO dt notation"""
    if split in s:
        n, s = s.split(split)
    else:
        n = 0
    if n == '':
        n = 0
    return int(n), s


def iso_hrs(s):
    """ Convert ISO dt notation to hours.
    Values except Weeks, Days, Hours ignored."""
    if s is None:
        return 0
    # Remove prefix
    s = s.split('P')[-1]
    # Step through letter dividers
    weeks, s = _iso_split(s, 'W')
    days, s = _iso_split(s, 'D')
    _, s = _iso_split(s, 'T')
    hours, s = _iso_split(s, 'H')
    # Convert all to hours
    return (weeks * 5 + days) * 8 + hours


@lru_cache(maxsize=None)  # Caching access to YT
def issue_times(issue):
    """ Return reverse-sorted by time list of issue spends, estimates, status and resolution changes"""
    sp = [{'date': dt.datetime.strptime(log.updatedAt, '%Y-%m-%dT%H:%M:%S.%f%z'),
           'by': log.updatedBy.display,
           'kind': field['field'].id,
           'value': field['to'] if field['field'].id in ['spent', 'estimation']
           else field['to'].key if field['to'] is not None else '',
           'from': field['from'] if field['field'].id in ['spent', 'estimation']
           else field['from'].key if field['from'] is not None else ''}
          for log in issue.changelog for field in log.fields
          if field['field'].id in ['spent', 'estimation', 'resolution', 'status']]
    sp.sort(key=lambda d: d['date'], reverse=True)
    return sp


@lru_cache(maxsize=None)  # Caching access to YT
def linked_issues(issue):
    def _accessible(someone):
        try:
            x = someone.summary is not None
        except Forbidden:
            x = False
        return x

    """ Return list of issue linked subtasks """
    return [link.object for link in issue.links
            if link.type.id == 'subtask' and
            dict(outward=link.type.inward, inward=link.type.outward)[link.direction] == 'Подзадача' and
            _accessible(link.object)]
