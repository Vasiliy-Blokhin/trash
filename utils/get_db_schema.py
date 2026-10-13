import re
from collections import defaultdict

from django.db import connection

from Site.agents.utils.build_where_condition import build_where_condition


_IDENTIFIER_RE = re.compile(r'^[A-Za-z_][A-Za-z0-9_$]{0,62}$')


def _is_valid_identifier(value) -> bool:
    return isinstance(value, str) and bool(_IDENTIFIER_RE.match(value))


def get_schema_mapping_context(db_schema='public', name='rel_db_struct', include_list=None):
    """
    Загружает таблицу маппинга и формирует иерархический справочник:
    - Схема (если таблица и поле пустые)
    - Таблица (если поле пустое)
    - Поле (если все заполнено)
    """
    # Имена схемы/таблицы подставляются в SQL как идентификаторы — проверяем их
    if not _is_valid_identifier(db_schema) or not _is_valid_identifier(name):
        return f"ОШИБКА: Некорректное имя таблицы маппинга '{db_schema}.{name}'."

    search_exp = build_where_condition(include_list or [])

    with connection.cursor() as cursor:
        # Проверяем существование таблицы маппинга
        # Колонки: schema_name, table_name, column_name, *_comment
        cursor.execute(
            """
            SELECT 1 FROM information_schema.tables
            WHERE table_schema = %s AND table_name = %s
            """,
            [db_schema, name],
        )
        if not cursor.fetchone():
            return f"ОШИБКА: Таблица маппинга '{db_schema}.{name}' не найдена."

        cursor.execute(f"""
            SELECT schema_name, table_name, column_name, schema_comment, table_comment, column_comment
            FROM "{db_schema}"."{name}"
            {search_exp}
            ORDER BY schema_name, table_name, column_name
        """)
        rows = cursor.fetchall()

    # Структура данных: { db_schema: { 'name': str, 'tables': { table: { 'name': str, 'cols': {col: name} } } } }
    schema_map = defaultdict(lambda: {'name': None, 'tables': defaultdict(lambda: {'name': None, 'cols': {}})})
    for sch, tbl, col, logic_sch, logic_tab, logic_fd in rows:
        # Без схемы строку невозможно привязать к иерархии
        if not sch:
            continue

        # Обработка уровня Схемы
        sch_data = schema_map[sch]
        if logic_sch:
            sch_data['name'] = logic_sch

        # Обработка уровня Таблицы
        if tbl:
            tbl_data = sch_data['tables'][tbl]
            if logic_tab:
                tbl_data['name'] = logic_tab

            # Обработка уровня Поля
            if col:
                tbl_data['cols'][col] = logic_fd

    if not schema_map:
        return "Справочник маппинга пуст."

    # Формирование текстового представления для промпта
    output_lines = [
        "### СПРАВОЧНИК МАППИНГА ИМЕН В БАЗЕ ДАННЫХ ###",
        "Правила использования:",
        "1. Для генерации SQL используй ТОЛЬКО физические имена (после стрелки).",
        "2. Логические имена (перед стрелкой) используй только для понимания запроса пользователя, не вставляй их в SQL.",
        "3. Обращай особое внимание на принадлежность поля к таблице, а таблицы к схеме.",
        "4. Формат физического имени в SQL: \"схема\".\"таблица\".\"поле\" или просто схема.таблица.поле",
        "",
    ]
    for sch_name, sch_data in schema_map.items():
        log_sch = sch_data['name'] or sch_name
        output_lines.append(f"[СХЕМА]{log_sch} -> {sch_name}")

        for tbl_name, tbl_data in sch_data['tables'].items():
            log_tbl = tbl_data['name'] or tbl_name
            # Отступ для таблицы
            output_lines.append(f"  └── [ТАБЛИЦА]{log_tbl} -> {sch_name}.{tbl_name}")

            for col_name, col_logic in tbl_data['cols'].items():
                # Отступ для поля
                output_lines.append(f"      ├── [ПОЛЕ]{col_logic or col_name} ->  {col_name}")
        output_lines.append("")  # Пустая строка между схемами

    return "\n".join(output_lines)
