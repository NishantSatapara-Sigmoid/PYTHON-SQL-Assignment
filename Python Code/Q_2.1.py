import xlrd
import psycopg2
from config import config
import pandas
import openpyxl

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
        query = "Insert into newTable (empno, ename, dname, compensation, months) values (%s,%s,%s,%s,%s)"
        for index, row in df.iterrows():
            cur.execute(query, (
            row['Employee Number'], row['Employee Name'], row['Department Name'], row['Compensation'],
            row['Total Months']))

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