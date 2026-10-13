from psycopg.rows import dict_row

from workspace.db_connection.manager.database import DatabaseManager
from workspace.db_connection.manager.config import DatabaseConfig



class DBWorker:
    def __init__(self):
        pass

    @staticmethod
    def run_psql(command, get=True):
        db = None
        try:
            db = DatabaseManager(DatabaseConfig())

            with db._get_connection() as conn:
                with conn.cursor(row_factory=dict_row) as cur:
                    cur.execute(query=command, prepare=True, binary=True)
                    result = cur.fetchall() if get else None
                conn.commit()
            return result
        finally:
            if db is not None:
                db.close()
