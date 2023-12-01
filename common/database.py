from sqlalchemy import create_engine
from sqlalchemy import text
from sqlalchemy import Engine
from urllib.parse import quote_plus as urlquote

def engine(db_name):
  '''
  '''
  db_user = ''
  db_password = ''
  db_host = ''
  db_port = 9030
  engine = create_engine(
      f'mysql+pymysql://{db_user}:{urlquote(db_password)}@{db_host}:{db_port}/{db_name}', echo=False)
  return engine

def execute(sql, engine: Engine):
    with engine.connect() as conn:
        result=conn.execute(text(sql))
        rows_affected = result.rowcount
        conn.commit()
        print('执行语句', sql)
    return rows_affected

