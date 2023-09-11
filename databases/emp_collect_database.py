import os
from contextlib import contextmanager

from sqlalchemy import func

from databases.database_connection import DatabaseConnection
from databases.models import Employee


@contextmanager
def database_session(session):
    try:
        yield session
    finally:
        session.close()


class EMPCollectDatabase(DatabaseConnection):
    def __init__(self):
        super().__init__('emp_collect')

    def is_high_band(self, email):
        with database_session(self.session) as session:
            result = session.query(Employee.band).filter(
                func.lower(Employee.email_primary_work) == func.lower(email)).first()
            if result:
                band_value = self.safe_string_to_int(result[0], default=0)
                min_ignore_band = int(os.getenv('QUARTERLY_ASSET_MIN_IGNORE_BAND'))
                if band_value >= min_ignore_band:
                    return True
                else:
                    return False
            else:
                return False

    def safe_string_to_int(self, value, default=None):
        try:
            if value is None:
                return default
            value = str(value).strip()
            if value == "":
                return default
            return int(value)
        except (ValueError, TypeError):
            return default
