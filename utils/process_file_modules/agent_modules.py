import os
import re
import ssl
import inspect
import functools

from mysite.logger import logger

# LOGGER
# -----------------------------------------------------------------------------
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
_THINK_END_RE = re.compile(r'</think>', flags=re.IGNORECASE)
_CODE_FENCE_RE = re.compile(r'```[ \t]*(?:python|json|markdown)?[ \t]*\n?', flags=re.IGNORECASE)
_LEADING_LANG_RE = re.compile(r'^\s*(?:python|json|markdown)[ \t]*\n', flags=re.IGNORECASE)


def clean_response(text, replace_quotes=True):
    if not text: return None
    if not isinstance(text, str):
        text = str(text)

    # Отбрасываем рассуждения модели: всё до последнего </think>
    last_think = None
    for last_think in _THINK_END_RE.finditer(text):
        pass
    if last_think is not None:
        text = text[last_think.end():].lstrip('\n')

    # Убираем markdown-ограждения кода и метку языка в начале ответа.
    # Слова python/json/markdown внутри текста не трогаем.
    text = _CODE_FENCE_RE.sub('', text)
    text = _LEADING_LANG_RE.sub('', text)
    if replace_quotes:
        text = text.replace("'", '"')
    return text

_FILE_NAME_FORBIDDEN_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

def _file_name_cleaner(text):
    if not text: return ''
    text = _FILE_NAME_FORBIDDEN_RE.sub("", str(text))
    text = re.sub(r"^[_\s.]+|[_\s.]+$", "", text)
    text = re.sub(r"_+", "_", text)
    return text.strip()
# -----------------------------------------------------------------------------

# REQUEST SSL
async def create_request_ssl(cert, key, verify):
    for path in (cert, key):
        if path and not os.path.isfile(path):
            raise FileNotFoundError(f'Файл сертификата не найден: {path}')
    ssl_context = ssl.create_default_context()
    if cert:
        ssl_context.load_cert_chain(cert, key)
    # Порядок важен: check_hostname нужно выключить до CERT_NONE
    ssl_context.check_hostname = bool(verify)
    ssl_context.verify_mode = ssl.CERT_REQUIRED if verify else ssl.CERT_NONE
    return ssl_context
# -----------------------------------------------------------------------------

# Exception handler
def handle_exceptions(func):
    if inspect.iscoroutinefunction(func):
        @functools.wraps(func)
        async def async_wrapper(*args, **kwargs):
            try:
                return await func(*args, **kwargs)
            except Exception as e:
                raise Exception(f"функция {func.__name__} --> {e}") from e
        return async_wrapper

    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            raise Exception(f"функция {func.__name__} --> {e}") from e
    return wrapper

# -----------------------------------------------------------------------------
AI_CONTENT = lambda user_answer='', file_content='': f'''
{"Запрос пользователя: " if user_answer else ""}
{user_answer}
{"Содержание документов: " if file_content else ""}
{file_content}
'''
SPLIT_SYMBS = '*|*'


def name_to_translit(text):
    if not text: return ''
    translit_text = "".join(
        TRANSLIT_DICT.get(char.lower(), char) for char in str(text)
    )
    return re.sub(r"\s+", "", translit_text).lower()[:75]

TRANSLIT_DICT = {
"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e",
"ё": "yo", "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k",
"л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r",
"с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "c",
"ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "",
"э": "e", "ю": "yu", "я": "ya", " ": "_", ",": "_", ".": "_",
"/": "_", "\\": "_", "-": "_", "'": "_", '"': "_"
}
