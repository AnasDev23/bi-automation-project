"""Local SQL Server settings loaded from environment variables."""

import os
from dotenv import load_dotenv

# Load values from a local .env file when present. Never commit that file.
load_dotenv()

SERVER = os.getenv("SQL_SERVER", r"localhost\SQLEXPRESS")
DATABASE = os.getenv("SQL_DATABASE", "ABC_Retail_DB")
ODBC_DRIVER = os.getenv("SQL_ODBC_DRIVER", "ODBC Driver 18 for SQL Server")
TRUSTED_CONNECTION = os.getenv("SQL_TRUSTED_CONNECTION", "yes")
ENCRYPT = os.getenv("SQL_ENCRYPT", "yes")
TRUST_SERVER_CERTIFICATE = os.getenv("SQL_TRUST_SERVER_CERTIFICATE", "yes")
