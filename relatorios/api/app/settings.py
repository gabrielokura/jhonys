import os
from dataclasses import dataclass


@dataclass(frozen=True)
class Settings:
    firebird_host: str = os.getenv("FIREBIRD_HOST", "firebird30")
    firebird_port: int = int(os.getenv("FIREBIRD_PORT", "3050"))
    firebird_database: str = os.getenv("FIREBIRD_DATABASE", "")
    firebird_user: str = os.getenv("FIREBIRD_USER", "SYSDBA")
    firebird_password: str = os.getenv("FIREBIRD_PASSWORD", "")
    firebird_charset: str = os.getenv("FIREBIRD_CHARSET", "UTF8")

    @property
    def firebird_dsn(self) -> str:
        return f"{self.firebird_host}/{self.firebird_port}:{self.firebird_database}"


settings = Settings()
