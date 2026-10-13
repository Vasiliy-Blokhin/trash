import json
import re
import ssl
import functools
import httpx

from Site.variables.url.url import AI_MODELS_DICT
from Site.modules.llm.get_llm import get_llm_settings
from Site.variables.consts.consts import GIGACHAT_CERT, GIGACHAT_KEY
from Site.modules.request.abstract_request import AsyncHttpRequest, \
    AbstractRequest
from Site.modules.text.clean_response import clean_response
from Site.api.xls_import_to_db.modules.params import \
    JSON_MAX_RETRIES, TEST
from Site.modules.vector.text_to_vector import get_embedding
from mysite.logger import logger


_JSON_FEEDBACK = (
    '\n\nВ прошлый раз ты ответил не в формате JSON. '
    'Повтори ответ СТРОГО в виде одного JSON-объекта, '
    'без пояснений.'
)
_RETRY_FEEDBACK = (
    '\nВ прошлый раз ты сформировал ответ не так как ожидалось, '
    'либо не отправил вовсе. Повторная попытка'
)
_CODE_FENCE_RE = re.compile(r'```(?:json)?', flags=re.IGNORECASE)


class AIWorker:
    def __init__(self, system_prompt='', content='', key=None):
        self.full_url = ''
        self.short_url = ''
        self.request_data = ''
        self.system_prompt = system_prompt
        self.content = content
        self.key = key

    def get_message_list(self, system_prompt=None, content=None):
        if system_prompt is None:
            system_prompt = self.system_prompt
        if content is None:
            content = self.content
        return [
            {"role": "system", "content": system_prompt or ''},
            {"role": "user", "content": content or ''}
        ]

    def _check_key(self):
        if not self.key:
            raise RuntimeError(
                'Модель не выбрана: вызовите check_ai() или передайте key'
            )

    def _apply_settings(self, settings, key, temp, top):
        full_url, short_url, request_data, error = settings
        if error or not full_url or not isinstance(request_data, dict):
            raise ValueError(
                f'Не удалось получить настройки модели "{key}": '
                f'{error or "пустой ответ"}'
            )
        self.full_url, self.short_url, self.request_data = \
            full_url, short_url, request_data
        # 0 — допустимое значение, поэтому сравнение с None
        if temp is not None: self.request_data['temperature'] = temp
        if top is not None: self.request_data['top_p'] = top

    async def get_ai_settings(
            self, key, message_list, temp=None, top=None, add_params=None):
        self._apply_settings(
            await get_llm_settings(
                key=key,
                message_list=message_list,
                params=add_params or {}
            ),
            key, temp, top
        )

    async def check_ai(self):
        # model_keys = ['gigachat__3_0', 'gigachat', 'qwen']
        model_keys = ['gigachat', 'gigachat__3_0', 'qwen']
        available = AI_MODELS_DICT().keys()

        for key in model_keys:
            if key not in available:
                logger.warning(f'check_ai: модель "{key}" не настроена')
                continue
            try:
                await self.get_ai_settings(
                        key=key,
                        message_list=self.get_message_list(
                            system_prompt='Проверка связи',
                            content=''
                        )
                    )
                if await self.request_ai():
                    self.key = key
                    logger.info(f'Выбран: "{self.key}"')
                    return True
            except Exception as e:
                logger.warning(f'check_ai: модель "{key}" недоступна: {e}')
        return False

    @staticmethod
    def _parse_answer(answer_data):
        if not isinstance(answer_data, dict):
            return None
        choices = answer_data.get('choices')
        if not choices:
            if answer_data.get('id', None):
                return answer_data
            return None
        if not isinstance(choices, list) or not isinstance(choices[0], dict):
            return None
        content = (choices[0].get('message') or {}).get('content')
        return clean_response(content) if content else None

    async def request_ai(self, files=None):
        if not self.full_url:
            logger.error('request_ai: не заданы настройки модели')
            return None
        request = AsyncHttpRequest()
        try:
            answer_data = await request.make_request(
                url=self.full_url,
                method='POST',
                data=self.request_data,
                verify=False,
                cert=GIGACHAT_CERT,
                key=GIGACHAT_KEY,
                files=files
            )
            return self._parse_answer(answer_data)
        except Exception as e:
            logger.error(f'Ошибка при формировании ответа от модели: {e}')
            return None

    @staticmethod
    def extract_json(text):
        if not text:
            return None
        if isinstance(text, (dict, list)):
            return text
        text = _CODE_FENCE_RE.sub('', str(text)).strip()
        # Быстрый путь: ответ целиком является JSON
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
        # Иначе ищем первый корректный JSON-объект/массив внутри текста
        decoder = json.JSONDecoder()
        for match in re.finditer(r'[\{\[]', text):
            try:
                result, _ = decoder.raw_decode(text, match.start())
                return result
            except json.JSONDecodeError:
                continue
        return None

    async def ask_ai_json(self, retries=JSON_MAX_RETRIES, temp=None, top=None):
        self._check_key()
        feedback = ''
        for attempt in range(max(0, retries) + 1):
            await self.get_ai_settings(
                key=self.key,
                message_list=self.get_message_list(
                    content=(self.content or '') + feedback
                ),
                temp=temp,
                top=top
            )
            answer = await self.request_ai()
            result = self.extract_json(answer)
            if result is not None:
                return result
            logger.info(
                f'ask_ai_json: попытка {attempt + 1} — невалидный JSON: '
                f'{str(answer)[:200]}'
            )
            feedback = _JSON_FEEDBACK
        return None

    async def ask_ai(self, files=None, temp=None, top=None, add_params=None):
        self._check_key()
        await self.get_ai_settings(
            key=self.key,
            message_list=self.get_message_list(),
            temp=temp, top=top, add_params=add_params
        )
        return await self.request_ai(files)

    @staticmethod
    def _strip_answer(answer):
        return answer.lstrip(' \n') if isinstance(answer, str) else answer

    async def get_ai_answer(
            self, content='', system_prompt='', temp=None, top=None,
            files=None, add_params=None
        ):
        self.content = content
        self.system_prompt = system_prompt

        answer = await self.ask_ai(
            files=files, temp=temp, top=top, add_params=add_params
        )

        if not answer:
            self.content = content + _RETRY_FEEDBACK
            answer = await self.ask_ai(
                files=files, temp=temp, top=top, add_params=add_params
            )
            if not answer: return ''

        return self._strip_answer(answer)

    async def get_ai_json_answer(
            self, content='', system_prompt='', temp=None, top=None
        ):
        self.content = content
        self.system_prompt = system_prompt

        answer = await self.ask_ai_json(temp=temp, top=top)

        # Пустой {} / [] — корректный JSON, повторяем только при None
        if answer is None:
            self.content = content + _RETRY_FEEDBACK
            answer = await self.ask_ai_json(temp=temp, top=top)
            if answer is None: return ''

        return answer

    @staticmethod
    async def create_vector(text):
        if not text or not str(text).strip():
            raise ValueError('create_vector: пустой текст')
        return await get_embedding(text)

    # sync methods ------------------------------------------------------------
    def sync_get_ai_settings(self, key, message_list, temp=None, top=None):
        self._apply_settings(
            self.sync_get_llm_settings(
                key=key,
                message_list=message_list
            ),
            key, temp, top
        )

    def sync_request_ai(self):
        if not self.full_url:
            logger.error('sync_request_ai: не заданы настройки модели')
            return None
        request = SyncHttpRequest()
        try:
            answer_data = request.make_request(
                url=self.full_url,
                method='POST',
                data=self.request_data,
                verify=False,
                cert=GIGACHAT_CERT,
                key=GIGACHAT_KEY,
            )
            return self._parse_answer(answer_data)
        except Exception as error:
            logger.error(
                f'Ошибка при формировании ответа от модели: {error}'
            )
            return None

    def sync_get_ai_answer(self, content='', system_prompt='', temp=None, top=None):
        self.content = content
        self.system_prompt = system_prompt

        answer = self.sync_ask_ai(temp=temp, top=top)

        return self._strip_answer(answer)

    def sync_ask_ai_json(self, retries=JSON_MAX_RETRIES, temp=None, top=None):
        self._check_key()
        feedback = ''
        for attempt in range(max(0, retries) + 1):
            self.sync_get_ai_settings(
                key=self.key,
                message_list=self.get_message_list(
                    content=(self.content or '') + feedback
                ),
                temp=temp,
                top=top
            )
            answer = self.sync_request_ai()
            result = self.extract_json(answer)
            if result is not None:
                return result
            logger.info(
                f'ask_ai_json: попытка {attempt + 1} — невалидный JSON: '
                f'{str(answer)[:200]}'
            )
            feedback = _JSON_FEEDBACK
        return None

    def sync_get_ai_json_answer(
            self, content='', system_prompt='', temp=None, top=None
        ):
        self.content = content
        self.system_prompt = system_prompt

        answer = self.sync_ask_ai_json(temp=temp, top=top)

        if answer is None:
            self.content = content + _RETRY_FEEDBACK
            answer = self.sync_ask_ai_json(temp=temp, top=top)
            if answer is None: return ''

        return answer

    def sync_ask_ai(self, temp=None, top=None):
        self._check_key()
        self.sync_get_ai_settings(
            key=self.key,
            message_list=self.get_message_list(),
            temp=temp,
            top=top
        )
        return self.sync_request_ai()


    @staticmethod
    def sync_get_llm_settings(key, message_list, params=None):
        if key not in AI_MODELS_DICT().keys():
            return None, None, None, 'Модель не найдена' # (4) el on error

        llm_settings = AI_MODELS_DICT(message_list, params or {})[key]

        return (
            llm_settings['full_url'],
            llm_settings['short_url'],
            llm_settings['request_data'],
            False # error
        )


class SyncHttpRequest(AbstractRequest):
    """
    Класс для выполнения синхронных HTTP-запросов.
    Этот класс реализует метод `make_request` для выполнения синхронных HTTP-запросов и обработки различных типов данных и методов.
    """
    def make_request(self, url, method="GET", headers=None, data=None, files=None, cert=GIGACHAT_CERT, key=GIGACHAT_KEY, verify=True,
                           cacert=None, to_json=True, timeout=900.0):
        if not url:
            raise ValueError('make_request: не задан url')
        method = (method or '').upper()
        if method not in ('GET', 'POST'):
            raise ValueError(f"Неподдерживаемый тип запроса: {method}")
        if cacert:
            verify = True

        context = self.sync_create_request_ssl(cert, key, bool(verify), cacert)

        response = None
        try:
            with httpx.Client(verify=context, timeout=timeout) as client:
                if method == 'GET':
                    response = client.get(url, headers=headers, params=data)
                elif files:
                    response = client.post(
                        url, headers=headers, data=data, files=files
                    )
                else:
                    response = client.post(
                        url, headers=headers, json=data
                    )
            response.raise_for_status()
            return response.json() if to_json else response
        except Exception as error:
            body = response.text[:2000] if response is not None else ''
            logger.error(f"Ошибка запроса: {error}.\nОтвет: {body}")
            return None

    @staticmethod
    @functools.lru_cache(maxsize=16)
    def sync_create_request_ssl(cert, key, verify, cacert=None):
        # Контекст кэшируется: загрузка цепочки сертификатов на каждый
        # запрос — лишняя работа с диском и криптографией
        ssl_context = ssl.create_default_context(cafile=cacert)
        if cert:
            ssl_context.load_cert_chain(cert, key)
        # Порядок важен: check_hostname нужно выключить до CERT_NONE
        ssl_context.check_hostname = verify
        ssl_context.verify_mode = ssl.CERT_REQUIRED if verify else ssl.CERT_NONE
        return ssl_context

