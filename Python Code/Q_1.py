#!/usr/bin/python
import psycopg2
from config import config
import xlsxwriter

class Q_1:
    def connect():
        conn = None
        try:
            params = config()
            print('Connecting to the PostgreSQL database...')
            conn = psycopg2.connect(**params)
            # create a cursor
            cur = conn.cursor()
            # execute a statement

            # 1) list employee NUMERICs, names and their managers
            query_1 = 'SELECT x.empno, x.ename , y.ename FROM emp as x INNER JOIN emp as y ON x.mgr = y.empno'
            cur.execute(query_1)
            table = cur.fetchall()
            workbook = xlsxwriter.Workbook('/Users/nishant/PycharmProjects/postgres/q_1.xlsx')
            sheet = workbook.add_worksheet()
            sheet.write('A1', 'Employee Number')
            sheet.write('B1', 'Employee Name')
            sheet.write('C1', 'Manager Name')
            r = 1
            for i in table:
                sheet.write(r, 0, i[0])
                sheet.write(r, 1, i[1])
                sheet.write(r, 2, i[2])
                r = r + 1
                print(i)

            # close the communication with the PostgreSQL
            cur.close()
            workbook.close()
        except (Exception, psycopg2.DatabaseError) as error:
            print(error)
        finally:
            if conn is not None:
                conn.close()
                print('Database connection closed.')


    if __name__ == '__main__':
        connect()

