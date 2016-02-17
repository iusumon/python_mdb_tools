# import pyodbc

# filename = r'/home/zahid/workspace/misc/Increment.mdb'

# connection = pyodbc.connect('DRIVER={Microsoft Access Driver (*.mdb)};DBQ='+filename)
# connection = pyodbc.connect('DRIVER={Easysoft ODBC-ACCESS}; MDBFILE='+filename)
# cursor = conn.cursor()
import adodbapi
conn = adodbapi.connect("ss")
cur = conn.execute()
