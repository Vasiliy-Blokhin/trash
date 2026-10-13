import sys
import os
import re
from datetime import datetime

from psycopg.rows import dict_row


current_dir = os.path.dirname(os.path.abspath(__file__))
while not os.path.exists(os.path.join(current_dir, 'Site')):
    parent_dir = os.path.dirname(current_dir)
    if parent_dir == current_dir:  # дошли до корня диска — проект не найден
        break
    current_dir = parent_dir
if current_dir not in sys.path:
    sys.path.append(current_dir)

from Site.utils.text.random import text_random
from mysite.logger import logger
from Site.variables.consts.consts import DATABASES_URI
from Site.agents.math.database import DatabaseManager
from Site.agents.math.config import DatabaseConfig
from Site.agents.process_docx_create.modules.params import (
    COMMAND_GET_DATA,
    COMMAND_INSERT_DATA,
    SCHEMA_NAME,
    SCHEMA_NAME_DOC,
    T_ALL_DOCUMENTS,
    T_STRUCT_TAG,
    T_SAVE_DATA,
    FILTER_RAVNO,
)


_COLUMN_NAME_RE = re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')
_NUMBER_RE = re.compile(r'-?\d+(?:[.,]\d+)?')


class DBWorker:
    """ Предварительно предназначен для работы с БД."""

    def __init__(self):
        pass

    # Основной метод для работы с postgres
    @staticmethod
    def run_psql(command, get=True):
        result = "0"
        if not command or not isinstance(command, str):
            logger.error(f'run_psql: пустая или некорректная команда: {command!r}')
            return result

        db = None
        try:
            db = DatabaseManager(DatabaseConfig(host=DATABASES_URI))
            with db._get_connection() as conn:
                try:
                    with conn.cursor(row_factory=dict_row) as cur:
                        cur.execute(query=command, binary=True)
                        # fetchall у запроса без результата (INSERT без RETURNING) падает
                        if get and cur.description is not None:
                            result = cur.fetchall()
                    conn.commit()
                except Exception:
                    conn.rollback()
                    raise
        except Exception as e:
            logger.error(f'run_psql: ошибка выполнения запроса: {e}\nЗапрос: {command[:500]}')
        finally:
            if db is not None:
                try:
                    db.close()
                except Exception as e:
                    logger.warning(f'run_psql: ошибка закрытия соединения: {e}')
        return result

    @staticmethod
    def _quote(value) -> str:
        """Значение -> безопасный SQL-литерал в одинарных кавычках."""
        return "'" + str(value).replace('\x00', '').replace("'", "''") + "'"

    @staticmethod
    def _require(value, name):
        if value is None or (isinstance(value, str) and not value.strip()):
            raise ValueError(f'Не задан обязательный параметр "{name}"')
        return value

    @classmethod
    def get_document(cls, id_doc):
        cls._require(id_doc, 'id_doc')
        return cls.run_psql(
            COMMAND_GET_DATA(
                SCHEMA_NAME, T_ALL_DOCUMENTS
            ) + FILTER_RAVNO('id', cls._quote(id_doc))
        )

    # Получение документов, в зависимости от его типа
    @classmethod
    def get_all_documents(cls):
        return cls.run_psql(
            COMMAND_GET_DATA(
                SCHEMA_NAME, T_ALL_DOCUMENTS
            )
        )

    # Получение структурных тегов, в зависимости от заданных тегов
    @classmethod
    def get_struct_tags(cls, document_data):
        if not isinstance(document_data, dict) or not document_data.get('id'):
            raise ValueError('get_struct_tags: ожидается словарь документа с ключом "id"')
        return cls.run_psql(
            COMMAND_GET_DATA(SCHEMA_NAME, T_STRUCT_TAG) + FILTER_RAVNO(
                "adl_id", cls._quote(document_data['id'])
            )
        )

    @staticmethod
    def convert_to_numbers(input_string):
        if not input_string:
            return []
        number_strings = _NUMBER_RE.findall(str(input_string))
        return [float(num_str.replace(',', '.')) for num_str in number_strings]

    @staticmethod
    def list_in_tuple_for_bd(data: list):
        if not isinstance(data, (list, tuple, str)):
            raise TypeError(f'list_in_tuple_for_bd: ожидается list/tuple/str, получен {type(data).__name__}')
        tuple_str = '('
        for i, el in enumerate(data):
            if el not in ['[', ']', ',', ' ']:
                if i < len(data) - 2:
                    tuple_str += el + ', '
                    continue
                tuple_str += el
        return tuple_str + ')'

    @classmethod
    def get_types_sample(cls, id_sample):
        cls._require(id_sample, 'id_sample')
        result = cls.run_psql(
            command=COMMAND_GET_DATA() + FILTER_RAVNO(col="id", data=cls._quote(id_sample))
        )
        if not isinstance(result, list) or not result:
            logger.warning(f'get_types_sample: образец "{id_sample}" не найден')
            return None
        return result[0]

    @classmethod
    def save_struct_tag_in_db(cls, document_id, name, file_link, context_tag):
        cls._require(document_id, 'document_id')
        comment = '!ТЕСТОВАЯ ЗАПИСЬ!'
        time_stump = cls.get_time_stump()
        return cls.run_psql(
            command=COMMAND_INSERT_DATA(
                SCHEMA_NAME, T_STRUCT_TAG,
                '("id", "key", "abbr", "_create", "_update", "_order", "comment", "name", "file_link", "context_tag", "adl_id")',
                (text_random(40), text_random(40), text_random(40), time_stump, time_stump, 1, comment, name, file_link, context_tag, document_id)
            ),
            get=False
        )

    @classmethod
    def save_document_in_db(
        cls, id_document, name,
        description, vector, full_prompt
    ):
        cls._require(id_document, 'id_document')
        comment = '!ТЕСТОВАЯ ЗАПИСЬ!'
        time_stump = cls.get_time_stump()
        return cls.run_psql(
            command=COMMAND_INSERT_DATA(
                SCHEMA_NAME, T_ALL_DOCUMENTS,
                '("id", "key", "abbr", "_create", "_update", "_order", "comment", "name", "description", "vector", "full_prompt")',
                (id_document, text_random(40), text_random(40), time_stump, time_stump, 1, comment, name,  description, vector, full_prompt)
            ),
            get=False
        )

    # Сохранение файла в БД (гигачат)
    @classmethod
    def save_file_in_gigachat_db(cls, data: dict, real_path):
        if not isinstance(data, dict):
            raise TypeError('save_file_in_gigachat_db: data должен быть словарём')
        if not real_path or not os.path.isfile(real_path):
            raise FileNotFoundError(f'save_file_in_gigachat_db: файл не найден: {real_path}')

        time_stump = cls.get_time_stump()
        data['_order'] = 0
        data['_create'] = time_stump
        data['_update'] = time_stump
        data['abbr'] = ''
        data['is_fake'] = False
        data['size'] = os.path.getsize(real_path)

        bad_keys = [k for k in data if not isinstance(k, str) or not _COLUMN_NAME_RE.match(k)]
        if bad_keys:
            raise ValueError(f'save_file_in_gigachat_db: некорректные имена колонок: {bad_keys}')

        # Имена колонок в двойных кавычках: ("a", "b", ...)
        col = '(' + ', '.join(f'"{k}"' for k in data) + ')'
        return cls.run_psql(
            get=False,
            command=COMMAND_INSERT_DATA(
                SCHEMA_NAME_DOC,
                T_SAVE_DATA,
                col,
                tuple(data.values())
            )
        )

    @staticmethod
    def get_time_stump() -> str:
        # Локальное время с реальным смещением часового пояса, формат
        # 'YYYY-MM-DD HH:MM:SS.ffffff+03' (минуты смещения — только если не 0)
        now = datetime.now().astimezone()
        offset = now.strftime('%z') or '+0000'
        time_zone = offset[:3] if offset[3:] == '00' else f'{offset[:3]}:{offset[3:]}'
        return now.strftime("%Y-%m-%d %H:%M:%S.%f") + time_zone
