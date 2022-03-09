#!/usr/bin/python
import psycopg2
from config import config
import xlsxwriter

class Q_2:
    def connect():
        conn = None
        try:
            params = config()
            print('Connecting to the PostgreSQL database...')
            conn = psycopg2.connect(**params)
            # create a cursor
            cur = conn.cursor()

            # execute a statement

            # 2) list the Total compensation  given till his/her last date or till now of all the employees till date
            cur.execute(
                'CREATE view temp AS SELECT empno, ename, dname, emp.deptno FROM emp INNER JOIN dept ON emp.deptno= dept.deptno')
            cur.execute(
                'CREATE view temp2 AS select temp.empno,temp.ename,temp.dname,temp.deptno,startdate,CASE WHEN enddate is NULL then CURRENT_DATE ELSE enddate END enddate,sal from temp INNER JOIN jobhist on temp.empno = jobhist.empno')
            query_2 = "SELECT empno,ename,dname,((DATE_PART('year', enddate::date) - DATE_PART('year', startdate::date)) * 12 + (DATE_PART('month', enddate::date) - DATE_PART('month', startdate::date)))*sal AS compansation, ((DATE_PART('year', enddate::date) - DATE_PART('year', startdate::date)) * 12 + (DATE_PART('month', enddate::date) - DATE_PART('month', startdate::date))) as Month from temp2"
            cur.execute(query_2)
            table = cur.fetchall()

            workbook = xlsxwriter.Workbook('q_2.xlsx')
            sheet = workbook.add_worksheet()

            sheet.write('A1', 'Employee Number')
            sheet.write('B1', 'Employee Name')
            sheet.write('C1', 'Department Name')
            sheet.write('D1', 'Compensation')
            sheet.write('E1', 'Total Months')

            row = 1
            for num, name, dname, comp, month in table:
                sheet.write(row, 0, num)
                sheet.write(row, 1, name)
                sheet.write(row, 2, dname)
                sheet.write(row, 3, comp)
                sheet.write(row, 4, month)
                row += 1
                print(num, name, dname, comp, month)

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


