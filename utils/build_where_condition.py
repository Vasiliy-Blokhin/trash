def _escape_literal(value) -> str:
    """Экранирует значение для подстановки в SQL-литерал в одинарных кавычках."""
    return str(value).replace('\x00', '').replace("'", "''")


def build_where_condition(data, column='schema_name'):
    if not data:
        return ""
    if isinstance(data, str):
        data = [data]

    conditions = []
    seen = set()
    for item in data:
        if item is None:
            continue
        item = str(item).strip()
        if not item or item in seen:
            continue
        seen.add(item)
        conditions.append(f"({column}='{_escape_literal(item)}')")

    if not conditions:
        return ""

    return "WHERE " + " OR ".join(conditions)
