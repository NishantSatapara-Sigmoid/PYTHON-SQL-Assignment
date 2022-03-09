import xlrd
import psycopg2
from config import config
import pandas
import openpyxl

class Q_2_1_:
    def connect():
        conn = None
        try:
            params = config()
            print('Connecting to the PostgreSQL database...')
            conn = psycopg2.connect(**params)
            # create a cursor
            cur = conn.cursor()
            # execute a statement
            df = pandas.read_excel('/Users/nishant/PycharmProjects/postgres/q_2.xlsx')
            q="CREATE TABLE newTable (empno NUMERIC(4) ,ename VARCHAR(10),dname VARCHAR(20),compensation NUMERIC(15,5),months NUMERIC(4));"
            cur.execute(q)
            query = "Insert into newTable (empno, ename, dname, compensation, months) values (%s,%s,%s,%s,%s)"
            for index, row in df.iterrows():
                cur.execute(query, (
                    row['Employee Number'], row['Employee Name'], row['Department Name'], row['Compensation'],
                    row['Total Months']))

            w="select * from newTable"
            cur.execute(w)
            table = cur.fetchall()
            for i in table:
                print(i)
            # close the communication with the PostgreSQL
            cur.close()

        except (Exception, psycopg2.DatabaseError) as error:
            print(error)
        finally:
            if conn is not None:
                conn.close()
                print('Database connection closed.')

    if __name__ == '__main__':
        connect()