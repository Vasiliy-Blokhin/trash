import json

from workspace.agent_modules.ai_worker import AgentAIWorker
from workspace.agent_modules.db_worker import AgentDBWorker
from workspace.agent_modules.docx_worker import DocxWorker
from workspace.agent_params.prompts import (
    PROMPT_TAGS_SELECTION
)



class AgentCheckDocument:
    def __init__(
            self,
            content=None,
            request_data=None,
            file_content=None,
            full_url=None,
            short_url=None,
            task_id=None
        ):
        self.content = content
        self.request_data = request_data
        self.file_content = file_content
        self.full_url = full_url
        self.short_url = short_url
        self.task_id = task_id

        # Сюда запрос писать
        self.user_prompt = 'Сформируй предложения в план комплектования '         'учебного центра по заявкам. Номер п/п - 1. Количество л/с в одном '         'потоке - 16, количество потоков в год - 4. Сроки обучения: '         '1-15.09, 1-15.10, 1-15.11, 1-15.12. Количество дней обучения по '         'срокам: УНВ РХБЗ - 4, 3, -, -; УНВ РЭБ - 4, 6, 8, -. Кому подаются '         'предложения: должность - Начальник оперативного управления штаба; '         'войсковая часть - 31600; воинское звание - полковник; инициал и '         'фамилия - А.Давыдов.'

        # Сюда название документа писать
        self.name = 'Предложения в план комплектования учебного центра в '         'следующем учебном году'

        # Сюда описание документа писать
        self.doc_description = 'документ с предложениями в план '         'комплектования учебного центра на следующий учебный год, '         'сформированный по заявкам органов военного управления (ОВУ) на '         'обучение личного состава. Содержит таблицу, в которой для каждой '         'учебной группы и каждого ОВУ указаны номер п/п, код учебной '         'группы, количество личного состава в одном потоке и количество '         'потоков в год в соответствии с требованиями приказа МО РФ «О сети '         'и емкости», а также количество дней обучения по каждому сроку '         'обучения. Число строк таблицы зависит от количества поданных '         'заявок. Документ подписывается должностным лицом, которому '         'подаются предложения.'

        self.ai_worker = AgentAIWorker()
        self.db_worker = AgentDBWorker()
        self.docx_worker = DocxWorker()

    async def __call__(self, *args, **kwds):
        try:
            # TODO get tags
            tags_dict = json.loads(
                open(
                    file='./in_data/tags.txt', mode='r', encoding='utf-8'
                ).read()
            )

            # TODO get context
            context = f'###Запрос пользователя:\n{self.user_prompt}' \
            f'###Содержание файла:\n{self.file_content}'

            # TODO AI create res
            self.tags_result = {}
            for tag, description in tags_dict.items():
                self.tags_result[tag] = await self.ai_worker.get_ai_answer(
                    content=context,
                    system_prompt=PROMPT_TAGS_SELECTION(description)
                )

            # TODO render docx and save
            self.docx_worker.process_document(
                input_path='./in_data/template.docx',
                output_path='./out_data/result.docx',
                text_processor=self._render
            )

            # TODO save embedding
            try:
                embedding = await self.ai_worker.create_vector(
                    self.name + self.doc_description
                )
                open(
                    file='out_data/embedding.txt', mode='w', encoding='utf-8'
                ).write(str(embedding))
            except Exception as e:
                # у AI_Manager нет модели эмбеддингов — документ всё равно
                # сохранён, эмбеддинг формируется на основном стенде
                return {
                    'success': True,
                    'result': f'success (без эмбеддинга: {e})',
                    'tags': self.tags_result
                }
            return {
                'success': True,
                'result': 'success',
                'tags': self.tags_result
            }
        except Exception as e:
            return {
                'success': False,
                'error': f'Ошибка в работе агента проверки шаблонов: {e}'
            }

    def _render(self, text, info):
        for tag, result in self.tags_result.items():
            text = text.replace('{' + '{ ' + tag +  ' }' + '}', result)
        return text