def _clean_identifier(part: str):
    part = part.strip().strip('"').strip()
    return part or None


def parse_db_physical_name(name_list=None):
    if name_list is None:
        return []
    if isinstance(name_list, str):
        name_list = [name_list]
    if not isinstance(name_list, (list, tuple, set)):
        raise TypeError('Преобразование схем данных завершилось с ошибкой: получен тип данных, отличный от строкового и массива')

    result = []
    for n in name_list:
        if not isinstance(n, str) or not n.strip():
            continue
        name_pieces = [_clean_identifier(p) for p in n.strip().split(".")]
        match len(name_pieces):
            case 1:
                schema_name, table_name = None, name_pieces[0]
            case 2 | 3:  # схема.таблица[.поле]
                schema_name, table_name = name_pieces[0], name_pieces[1]
            case _:
                continue
        if table_name:
            result.append({"schema_name": schema_name, "table_name": table_name})
    return result
