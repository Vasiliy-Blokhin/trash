import datetime
import logging
import re
import functools


# LOGGER
# -----------------------------------------------------------------------------
level = 'DEBUG'

logger = logging.getLogger('gr13_pr1')
logger.setLevel(level)

console_handler = logging.StreamHandler()
console_handler.setLevel(level)

class CustomFormatter(logging.Formatter):
    def formatTime(self, record, datefmt=None):
        ct = datetime.datetime.fromtimestamp(record.created)
        if datefmt:
            s = ct.strftime(datefmt)
        else:
            try:
                s = ct.isoformat(timespec='milliseconds')
            except TypeError:
                s = ct.isoformat()
        return s[:23]

formatter = CustomFormatter(
    '%(asctime)s %(levelname)-8s %(message)s', datefmt='%Y-%m-%d %H:%M:%S,%f'
)

console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

def ld(message):
    logger.debug(message)

def li(message):
    logger.info(message)

def le(message):
    logger.error(message)

def lw(message):
    logger.warning(message)
# -----------------------------------------------------------------------------

# CLEAN RESPONSE
def clean_response(text):
    if not text: return None
    
    code_list = ['python', 'json', 'markdown']
    for code in code_list:
        if '</think>' in text:
            positions = [m.start() for m in re.finditer(
                r'</think>', text, flags=re.IGNORECASE
            )]
            if positions:
                last_position = positions[-1]
            text = text[last_position:]
            text = re.sub( r'^</think>\n?', '', text, flags=re.IGNORECASE)
        text = re.sub(rf'```{code}\n?', '', text)
        text = re.sub(rf'{code}\n?', '', text)
        text = re.sub(r'```\n?', '', text)
        text = re.sub( rf'^{code}\n?', '', text, flags=re.IGNORECASE)
        text = re.sub( rf'```', '', text, flags=re.IGNORECASE)
        text = text.replace("'", '"')
    return text
# -----------------------------------------------------------------------------

# REQUEST SSL
import ssl


async def create_request_ssl(cert,key, verify):
    ssl_context = ssl.create_default_context()
    ssl_context.load_cert_chain(cert,key)
    ssl_context.check_hostname = verify
    ssl_context.verify_mode = ssl.CERT_NONE
    return ssl_context
# -----------------------------------------------------------------------------

# Exception handler
def handle_exceptions(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            raise Exception(f"функция {func.__name__} --> {e}") from e
    return wrapper
