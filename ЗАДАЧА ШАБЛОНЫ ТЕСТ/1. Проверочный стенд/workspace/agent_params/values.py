TRANSLIT_DICT = {
"а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e",
"ё": "yo", "ж": "zh", "з": "z", "и": "i", "й": "y", "к": "k",
"л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r",
"с": "s", "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "c",
"ч": "ch", "ш": "sh", "щ": "sch", "ъ": "", "ы": "y", "ь": "",
"э": "e", "ю": "yu", "я": "ya", " ": "_", ",": "_", ".": "_",
"/": "_", "\\": "_", "-": "_", "'": "_", '"': "_"
}

ALLOWED_CHARS = (
    'abcdefghijklmnopqrstuvwxyz'
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
)

SCHEMA_NAME = 'pass'
TABLE_NAME = 'pass'
COMMAND_INSERT_DATA = lambda schema, table, keys, values: f'''
INSERT INTO {schema}."{table}" {keys} VALUES {values}
'''

# TODO change to url with docs
PRIKAZ_22_PATH = './in_data/Приказ_МО_РФ__22_от_17.01.2022.docx'

SAMPLE_SUCCESS_GR13_PR_1 = './in_data/sample_gr13_pr1_success.docx'
SAMPLE_FAILED_GR13_PR_1 = './in_data/sample_gr13_pr1_failed.docx'
SAMPLE_SUCCESS_GR13_PR_2 = './in_data/sample_gr13_pr2_success.docx'
SAMPLE_FAILED_GR13_PR_2 = './in_data/sample_gr13_pr2_failed.docx'
# TODO when send in POTOK --> uploads/sample
SAMPLE_SUCCESS = lambda type_document: SAMPLE_SUCCESS_GR13_PR_1 if \
    type_document == TYPE_VOZMOZHNOST else SAMPLE_SUCCESS_GR13_PR_2
SAMPLE_FAILED = lambda type_document: SAMPLE_FAILED_GR13_PR_1 if \
    type_document == TYPE_VOZMOZHNOST else SAMPLE_FAILED_GR13_PR_2

# TODO when send in POTOK --> uploads/doc_create_result
RESULT_PATH =  './out_data/'

TYPE_VOZMOZHNOST = 'zakluchenie_o_vozmozhnosti_otkrytogo_opublokovania'
TYPE_NALICHII = 'zakluchenie_o_nalichii_otsutstvii_svedeniy'

SUCCESS_MESSAGE = lambda check_result: f'''
Результат работы: 
В документе {
    "найдены" if check_result.get(
        'check_flag', None
    ) else "не найдены"
} критерии соответствующие пунктам Перечня.

Сформированы следующие документы:
1. заключение о возможности открытого опубликования материалов;
2. Формирование заключения о наличии сведений составляющих служебную тайну.
'''

