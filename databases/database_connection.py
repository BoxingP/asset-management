import os
from urllib.parse import quote

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker

Session = sessionmaker()


class DatabaseConnection(object):
    def __init__(self, database_name: str):
        self.database_name = database_name.upper()
        self.session = self._create_session()

    def _create_session(self):
        adapter = os.getenv(f'{self.database_name}_ADAPTER')
        host = os.getenv(f'{self.database_name}_HOST')
        port = os.getenv(f'{self.database_name}_PORT')
        user = os.getenv(f'{self.database_name}_USER')
        password = os.getenv(f'{self.database_name}_PASSWORD')
        database_str = os.getenv(f'{self.database_name}_DATABASE_STR')
        db_uri = f'{adapter}://{user}:%s@{host}:{port}/{database_str}' % quote(password)
        engine = create_engine(db_uri, echo=False)
        Session.configure(bind=engine)
        return Session()
