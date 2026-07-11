from contextlib import contextmanager
from typing import Iterator

from firebird.driver import Connection, connect

from .settings import settings


@contextmanager
def firebird_connection() -> Iterator[Connection]:
    conn = connect(
        settings.firebird_dsn,
        user=settings.firebird_user,
        password=settings.firebird_password,
        charset=settings.firebird_charset,
    )
    try:
        yield conn
    finally:
        conn.close()


def ping_firebird() -> int:
    with firebird_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("select 1 from rdb$database")
        row = cursor.fetchone()
        return int(row[0])
