import os
import configparser
import pymssql

class db():
    def __init__(self):
        proDir = os.getcwd()
        configPath = os.path.join(proDir, "sql_information.ini")
        config = configparser.ConfigParser()
        config.read(configPath)

        SQLpath = config['SQL']

        self.server_230 = SQLpath['server_230']
        self.user_230 = SQLpath['user_230']
        self.password_230 = SQLpath['password_230']
        self.database_230 = SQLpath['database_230']

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def conn(self, D):
        if D == 230:
            self.connect = pymssql.connect(server=self.server_230 , user=self.user_230 ,  password=self.password_230 , database=self.database_230 )
        
        self.cursor = self.connect.cursor()


    def test_conn(self, D):
        try:
            self.conn(D)
            self.cursor.execute("SELECT 1")
            result = self.cursor.fetchone()
            if result[0] == 1:
                return True
            else:
                return False
        except Exception as e:
            print(f"Database connection failed with error: {e}")
            return False
        finally:
            self.close()

    def get_datatable(self, sql, D):
        self.conn(D)
        self.cursor.execute(sql)
        x = self.cursor.fetchall()
        self.close()
        return x

    def run_cmd(self, sql, D):
        self.conn(D)
        if hasattr(self, 'cursor'):
            self.cursor.execute(sql)
        if hasattr(self, 'connect'):
            self.connect.commit()
        self.close()

    def close(self):
        if hasattr(self, 'cursor'):
            self.cursor.close()
        if hasattr(self, 'connect'):
            self.connect.close()



