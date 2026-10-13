from dataclasses import dataclass

from workspace.workspace_params.settings import DB_CONFIG


@dataclass
class DatabaseConfig:
    host: str = DB_CONFIG['host']
    port: int = DB_CONFIG['port']
    database: str = DB_CONFIG['database']
    user: str = DB_CONFIG['user']
    password: str = DB_CONFIG['password']
    min_connections: int = DB_CONFIG['min_connections']
    max_connections: int = DB_CONFIG['max_connections']
    statement_timeout: str = DB_CONFIG['statement_timeout']

    @property
    def url(self) -> str:
        return (
            f"host={self.host} port={self.port} "
            f"dbname={self.database} user={self.user} "
            f"password={self.password}"
        )

    @property
    def connection_kwargs(self) -> dict:
        return {
            "options": f"-c statement_timeout={self.statement_timeout}"
        }


