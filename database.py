from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

DATABASE_URL = "sqlite:///schedule.db"

engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Database:
    def __init__(self):
        self.session = SessionLocal()

    def init_db(self):
        Base.metadata.create_all(bind=engine)

    def drop_db(self):
        Base.metadata.drop_all(bind=engine)

    def recreate_db(self):
        self.drop_db()
        print('db dropped!')
        self.init_db()

    def close(self):
        self.session.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def add(self, obj):
        self.session.add(obj)

    def commit(self):
        self.session.commit()