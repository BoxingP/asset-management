from sqlalchemy import Column, Integer, VARCHAR, NVARCHAR
from sqlalchemy.orm import declarative_base

Base = declarative_base()


class ChinaVIP(Base):
    __tablename__ = 'china_vip'

    id = Column(Integer, primary_key=True)
    email = Column(VARCHAR, nullable=False)


class Employee(Base):
    __tablename__ = 'V_EMPLOYEE_ITAsset'

    employee_id = Column(VARCHAR(20), primary_key=True)
    worker_name = Column(NVARCHAR(200))
    email_primary_work = Column(NVARCHAR(100))
    band = Column(NVARCHAR(200))
