from contextlib import contextmanager

from databases.database_connection import DatabaseConnection
from databases.models import ChinaVIP


@contextmanager
def database_session(session):
    try:
        yield session
    finally:
        session.close()


class EMPInfoDatabase(DatabaseConnection):
    def __init__(self):
        super().__init__('emp_info')

    def is_china_vip(self, email):
        with database_session(self.session) as session:
            query = session.query(ChinaVIP).filter(ChinaVIP.email.ilike(email))
            result = query.first()
            if result:
                return True
            else:
                return False
