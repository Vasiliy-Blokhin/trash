import json
import re
import asyncio
from pptx import Presentation
import pandas as pd

from workspace.ai_connection.connect_llm import AsyncHttpRequest, \
    SyncHttpRequest
from workspace.workspace_params.modules import clean_response, li, le
from workspace.workspace_params.settings import (
    QWEN_SETTINGS, EMBEDDING_SETTINGS, GIGACHAT_CERT, GIGACHAT_KEY,
    TOKENIZER_SETTINGS, OCR_SERVICE_URL, LLM_HEADERS, LLM_TIMEOUT, LLM_RETRIES
)


class AIWorker:
    def __init__(self, system_prompt='', content=''):
        self.full_url = ''
        self.short_url = ''
        self.request_data = ''
        self.system_prompt = system_prompt
        self.content = content

    def get_message_list(self, system_prompt=None, content=None):
        if not system_prompt:
            system_prompt = self.system_prompt
        if not content:
            content = self.content
        return [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content}
        ]

    async def get_ai_settings(self, message_list):
        self.full_url, self.short_url, self.request_data = \
            (
                QWEN_SETTINGS()['full_url'],
                QWEN_SETTINGS()['short_url'],
                QWEN_SETTINGS(message_list)['request_data']
            )
    
    async def get_embedding_settings(self, text):
        self.full_url, self.short_url, self.request_data = \
            (
                EMBEDDING_SETTINGS()['full_url'],
                EMBEDDING_SETTINGS()['short_url'],
                EMBEDDING_SETTINGS(text)['request_data']
            )
    
    async def get_calc_tokens_settings(self, text):
        self.full_url, self.short_url, self.request_data = \
            (
                TOKENIZER_SETTINGS()['full_url'],
                TOKENIZER_SETTINGS(text)['request_data']
            )

    async def request_ai(self):
        # при недоступности сервиса (перезапуск AI_Manager) повторяем запрос
        for attempt in range(LLM_RETRIES):
            answer = await self._request_ai_once()
            if answer is not None or attempt == LLM_RETRIES - 1:
                return answer
            await asyncio.sleep(30)

    async def _request_ai_once(self):
        request = AsyncHttpRequest()
        try:            
            answer_data = await request.make_request(
                url=self.full_url,
                method='POST',
                headers=LLM_HEADERS,
                data=self.request_data,
                verify=False,
                cert=GIGACHAT_CERT,
                key=GIGACHAT_KEY,
                timeout=LLM_TIMEOUT
            )
            choices = answer_data.get('choices', None)
            if not choices:
                if answer_data.get('id', None):
                    return answer_data
                return None
            return clean_response(
                choices[0].get('message', {}).get('content', None)
            )
        except Exception as e:
            le(f'Ошибка при формировании ответа от модели: {e}')
            return None
    
    async def ask_ai_json(self, retries=2):
        feedback = ''
        answer = None
        for attempt in range(retries + 1):
            self.content += feedback
            await self.get_ai_settings(
                message_list=self.get_message_list()
            )

            answer = await self.request_ai()
            answer = clean_response(answer)
            answer = '{' + answer.split('{')[-1]
            result = self._extract_json(answer)
            if result is not None:
                return result
            li(
                f'ask_ai_json: попытка {attempt + 1} — невалидный JSON: '
                f'{str(answer)[:200]}'
            )
            feedback = (
                '\n\nВ прошлый раз ты ответил не в формате JSON. '
                'Повтори ответ СТРОГО в виде одного JSON-объекта, '
                'без пояснений.'
            )
        return None
    
    async def ask_ai(self):
        await self.get_ai_settings(
            message_list=self.get_message_list()
        )
        return await self.request_ai()

    async def get_embedding(self, text):
        await self.get_embedding_settings(text)
        result = await self.request_ai()

        if result["data"][0]["embedding"]:
            return result["data"][0]["embedding"]
        else:
            raise ConnectionError(result["error"])
    
    async def calc_tokens(self, text):
        await self.get_embedding_settings(text)
        result = await self.request_ai()

        if result:
            return result["usage"]["total_tokens"]
        else:
            raise ConnectionError(result["error"])
        
    @staticmethod
    def _extract_json(text):
        if not text:
            return None
        text = re.sub(r'```(?:json)?', '', text)
        return json.loads(text)

        
    async def get_ai_answer(self, content='', system_prompt=''):
        answer = None
        self.content = content
        self.system_prompt = system_prompt
        answer = await self.ask_ai()

        if not answer:
            self.content = content + '\nВ прошлый раз ты сформировал ответ не \
                так как ожидалось, либо не отправил вовсе. Повторная попытка'
            answer = await self.ask_ai()
            if not answer: return ''
        
        while answer[0] == ' ' or answer[0] == '\n':
            answer = answer[1:]
        return answer
    
    async def get_ai_json_answer(self, content='', system_prompt=''):
        answer = None
        self.content = content
        self.system_prompt = system_prompt
        answer = await self.ask_ai_json()

        if not answer:
            self.content = content + '\nВ прошлый раз ты сформировал ответ не \
                так как ожидалось, либо не отправил вовсе. Повторная попытка'
            answer = await self.ask_ai_json()
            if not answer: return ''
        
        return answer
    
    async def create_vector(self, text):
        return await self.get_embedding(text)

    async def get_calc_tokens(self, text):
        return await self.calc_tokens(text)
    
    async def extract_text_from_jpeg(jpg_path):
        """ Для '.pdf' файлов использовать ключ 'pdf'. Для документов формата 
        '.xlsx', '.docx', '.pptx' сервис OCR не требуется (извлечение текста
        происходит одноименными библиотеками, кроме '.xlsx' --> pandas).
        Для '.docx' также можете использовать метод 'find_context_document' из
        'docx_worker'
        """
        request = AsyncHttpRequest()
        try:
            files = [
                ('file', (jpg_path, open(jpg_path, 'rb')))
            ]            
            answer_data = await request.make_request(
                url=OCR_SERVICE_URL("image"),
                method='POST',
                files=files,
                verify=False,
                cert=GIGACHAT_CERT,
                key=GIGACHAT_KEY
            )
            return json.loads(answer_data.text).get("text", "")
        except Exception as e:
            le(f"Ошибка при OCR обработке файла: {e}")
            return ''
        
    def extract_text_from_xlsx(input_path):
        """
        Читает текст из существующего файла DOCX и возвращает его, включая 
        текст из таблиц.
        """
        try:
            xls = pd.ExcelFile(input_path)
            extracted_text = ""

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

                sheet_text = "\n".join(["\t".join(
                    map(str, row)
                ) for row in df.values])
                extracted_text += f"\nЛист: {sheet_name}\n{sheet_text}"
            xls.close()
            return extracted_text.strip()

        except Exception as e:
            return f"Произошла ошибка: {str(e)}"


    def extract_text_from_pptx(input_path, extension):
        if extension == ".ppt":
            return """
                ###ЗАДАЧА:
                Передай ***ОТВЕТ ПОЛЬЗОВАТЕЛЮ*** без искажений! Тебе \
                ЗАПРЕЩЕНО проводить анализ текста, документа и тому подобное, \
                ТОЛЬКО ПЕРЕДАЧА ***ОТВЕТ ПОЛЬЗОВАТЕЛЮ***.

                ###ОТВЕТ ПОЛЬЗОВАТЕЛЮ: \
                # (всё что ниже это текст для пользователя!)

                Извлечение текста из "*.ppt" не поддерживается. \
                Сконвертируйте в формат "*.pptx" самостоятельно и повторите \
                попытку.

                ###Для конвертации необходимо выполнить следующие действия:
                ____________________________________________________________

                1. Откройте презентацию формата "*.ppt";
                2. Нажмите на панели вкладку "Файл";
                3. Перейдите во вкладку "Экспорт";
                4. Выберите "Изменить тип файла";
                5. Выберите "Презентация (.pptx)".
                6. Повторите попытку отправки документа для краткого анализа.
            """ 
        # В потоке извлеченный текст возвращается пользователю после обработки
        # ИИ в зависимости от запроса.

        prs = Presentation(input_path)
        text_runs = []

        for slide in prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        text_runs.append(run.text)

        return "\n".join(text_runs)
    
    # sync methods ------------------------------------------------------------

    # При работе с docx_worker.process_document может возникнуть конфликт с 
    # асинхронными запросами, поэтому внутри работы процессора разрешается 
    # использовать синхронный запрос, при этом итоговый метод должен 
    # оборачиваться в asgiref --> sync --> sync_to_async (примеры использования
    # есть в теле агента)
    def sync_get_ai_settings(self, key, message_list, temp=None, top=None):
        self.full_url, self.short_url, self.request_data, _ = \
            self.sync_get_llm_settings(
                key=key,
                message_list=message_list
            )
        if temp: self.request_data['temperature'] = temp
        if top: self.request_data['top_p'] = top

    def sync_request_ai(self):
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

            choices = answer_data.get('choices')
            if isinstance(choices, list):
                first_choice = choices[0]
                
            return clean_response(
                first_choice.get('message', {}).get('content')
            )
        except Exception as error:
            le(
                f'Ошибка при формировании ответа от модели: {error}'
            )
        
    def sync_get_ai_answer(self, content='', system_prompt='', temp=None, top=None):
        self.content = content
        self.system_prompt = system_prompt
        
        answer = self.sync_ask_ai(temp=temp, top=top)
        
        return answer
    
    def sync_ask_ai(self, temp=None, top=None):
        self.sync_get_ai_settings(
            key=self.key,
            message_list=self.get_message_list(),
            temp=temp,
            top=top
        )
        return self.sync_request_ai()
    
    def sync_get_llm_settings(self, message_list):
        self.full_url, self.short_url, self.request_data = \
            (
                QWEN_SETTINGS()['full_url'],
                QWEN_SETTINGS()['short_url'],
                QWEN_SETTINGS(message_list)['request_data']
            )
    # -------------------------------------------------------------------------
    