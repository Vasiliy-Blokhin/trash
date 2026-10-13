from contextlib import contextmanager
from typing import Optional

import psycopg
from psycopg_pool import ConnectionPool

from workspace.workspace_params.modules import le, li
from workspace.db_connection.manager.config import DatabaseConfig


class DatabaseManager:
    def __init__(self, config: DatabaseConfig):
        try:
            self.config = config
            self.pool: Optional[ConnectionPool] = None
            self._init_pool()
        except Exception as error:
            le(f"Ошибка инициализации менеджера БД: {error}")

    def _init_pool(self) -> None:
        try:
            self.pool = ConnectionPool(
                self.config.url,
                min_size=self.config.min_connections,
                max_size=self.config.max_connections,
                kwargs=self.config.connection_kwargs,
            )
            li(
                "Пул соединений инициализирован: "
                f"min={self.config.min_connections}, "
                f"max={self.config.max_connections}"
            )
        except psycopg.OperationalError as e:
            le(f"Ошибка инициализации пула соединений: {e}")

    @contextmanager
    def _get_connection(self):
        try:
            if not self.pool:
                raise RuntimeError("Пул соединений не инициализирован")

            with self.pool.connection() as conn:
                yield conn
        except Exception as error:
            le(f"Ошибка инициализации пула соединений: {error}")

    def close(self) -> None:
        try:
            if self.pool:
                self.pool.close()
                self.pool = None
                li("Пул соединений закрыт")
        except Exception as error:
            le(f'Ошибка при закрытии пула: {error}')

    def check_connection(self) -> bool:
        try:
            with self._get_connection() as conn:
                with conn.cursor() as cur:
                    cur.execute("SELECT 1")
                    return True
        except Exception as e:
            le(f"Ошибка подключения к БД: {e}")
            return False
