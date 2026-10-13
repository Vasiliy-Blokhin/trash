from datetime import datetime, date

from mysite.logger import logger


def for_date(o, iso=True):
    if o is None:
        return ''
    try:
        return o.strftime("%Y-%m-%d" if iso else "%d.%m.%Y")
    except Exception as e:
        logger.error(f'for_date: не удалось преобразовать {o!r}: {e}')
        return ''

def for_datetime(o, iso=True):
    if o is None:
        return ''
    try:
        return o.isoformat() if iso else o.strftime("%H:%M:%S %d.%m.%Y")
    except Exception as e:
        logger.error(f'for_datetime: не удалось преобразовать {o!r}: {e}')
        return ''

def my_convert_date(o, iso=True):
    # datetime — подкласс date, поэтому проверяется первым
    if isinstance(o, datetime):
        return for_datetime(o, iso)
    if isinstance(o, date):
        return for_date(o, iso)
    return ''
