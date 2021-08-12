import psycopg2
from openpyxl.workbook import Workbook
import pandas as pd
# write this line at top
#importing the required libraries

class Total_compensation:
    def compensation(self):
        #connecting the python file with the database
        try:
            connection = psycopg2.connect(
                host="localhost",
                database="PythonANDSQL",
                user="postgres",
                password="1234")
            cursor_object = connection.cursor()
            # write query in structured way(in kind of indented way)
            query_command = """select emp.ename, emp.empno, dept.dname, (case when enddate is not null then ((enddate-startdate+1)/30)*(jobhist.sal) else ((current_date-startdate+1)/30)*(jobhist.sal) end)as Total_Compensation,
(case when enddate is not null then ((enddate-startdate+1)/30) else ((current_date-startdate+1)/30) end)as Months_Spent from jobhist, dept, emp 
where jobhist.deptno=dept.deptno and jobhist.empno=emp.empno"""
            # sql command to show the desired data

            cursor_object.execute(query_command)
            
            # By using copy command in postgreSQL we can save result into .xlsx in fewer lines of code
            columns = [desc[0] for desc in cursor_object.description]
            data = cursor_object.fetchall()
            df = pd.DataFrame(list(data), columns=columns)

            writer = pd.ExcelWriter('number2.py.xlsx')
            df.to_excel(writer, sheet_name='bar')
            writer.save()

        except Exception as e:
            print("The Program has not run successfully", e)
            # Run if the program has any exceptions
        finally:
            # this will run in all test cases
            if connection is not None:
                cursor_object.close()
                connection.close()
            # Closing the connections created after the program has run

if __name__=='__main__':
    connection = None
    cursor_object = None
    comp = Total_compensation()         #Create a object of Total_Compensation class
    comp.compensation()
