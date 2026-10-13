import re

from workspace.db_connection.db_worker import DBWorker
from workspace.agent_params.modules import get_time_stump
from workspace.agent_params.values import (
    COMMAND_INSERT_DATA, SCHEMA_NAME, TABLE_NAME
)


class AgentDBWorker(DBWorker):
    def __init__(self):
        super().__init__()
    
    @classmethod
    def save_file_in_gigachat_db(self, data: dict):

        data['_order'] = 0
        data['_create'] = get_time_stump()
        data['_update'] = get_time_stump()
        data['abbr'] = ''
        data['is_fake'] = False
        keys = tuple(data.keys())
        values = tuple(data.values())

        col = ''
        for el in (str(keys)):
            if el in ["'"]:
                col += '"'
                continue
            col += el
        self.run_psql(
            flag_gigachat=True,
            get=False,
            command=COMMAND_INSERT_DATA(
                    SCHEMA_NAME,
                    TABLE_NAME,
                    col,
                    values
                )
            )
        
    @staticmethod
    def convert_to_numbers(input_string):
        number_strings = re.findall(r"-?\d+.\d+|-?\d+", input_string)
        return [float(num_str) for num_str in number_strings]
