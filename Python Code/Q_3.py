#!/usr/bin/python
import psycopg2
from config import config
import xlsxwriter

def connect():
    conn = None
    try:
        params = config()
        print('Connecting to the PostgreSQL database...')
        conn = psycopg2.connect(**params)
        # create a cursor
        cur = conn.cursor()

        # execute a statement

        #3)list total compensation given at Department level till date. Columns: Dept No, Dept,Name, Compensation

        cur.execute(
            'CREATE view temp AS SELECT empno, ename, dname, emp.deptno FROM emp INNER JOIN dept ON emp.deptno= dept.deptno')

        cur.execute(
            'CREATE view temp2 AS select temp.empno,temp.ename,temp.dname,temp.deptno,startdate,CASE WHEN enddate is NULL then CURRENT_DATE ELSE enddate END enddate,sal from temp INNER JOIN jobhist on temp.empno = jobhist.empno')

        query_3 = "SELECT dname,deptno,SUM(((DATE_PART('year', enddate::date) - DATE_PART('year', startdate::date)) * 12 + (DATE_PART('month', enddate::date) - DATE_PART('month', startdate::date)))*sal) AS compansation from temp2 group by dname,deptno"
        cur.execute(query_3)
        table = cur.fetchall()


        workbook = xlsxwriter.Workbook('q_2.xlsx')
        sheet = workbook.add_worksheet()

        sheet.write('A1', 'Department name')
        sheet.write('B1', 'Department Number')
        sheet.write('C1', 'Compensation')

        row = 1
        for name, number, comp in table:
            sheet.write(row,0, name)
            sheet.write(row,1, number)
            sheet.write(row,2, comp)
            row += 1
            print(name,number,comp)

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
