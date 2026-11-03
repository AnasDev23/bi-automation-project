"""SQL Server connection helper."""

from sqlalchemy import create_engine
from urllib.parse import quote_plus

from config.settings import (
    SERVER,
    DATABASE,
    ODBC_DRIVER,
    TRUSTED_CONNECTION,
    ENCRYPT,
    TRUST_SERVER_CERTIFICATE,
)


def create_database_engine():
    """Create a SQLAlchemy engine using Windows integrated authentication."""
    connection_string = (
        f"DRIVER={{{ODBC_DRIVER}}};"
        f"SERVER={SERVER};"
        f"DATABASE={DATABASE};"
        f"Trusted_Connection={TRUSTED_CONNECTION};"
        f"Encrypt={ENCRYPT};"
        f"TrustServerCertificate={TRUST_SERVER_CERTIFICATE};"
    )

    connection_url = (
        "mssql+pyodbc:///?odbc_connect="
        + quote_plus(connection_string)
    )

    return create_engine(connection_url)
