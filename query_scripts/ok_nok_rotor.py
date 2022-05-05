from datetime import datetime
import pyodbc
import pandas as pd
import time

conn_str = ( 
    r'Driver={ODBC Driver 17 for SQL Server};'
    r'Server=CZE2-SV00052\CZE2PMS1;'
    r'Database=SITMesDB;'
    r'Trusted_Connection=yes;'
)
cnxn = pyodbc.connect(conn_str)
cursor = cnxn.cursor()
script = """
SELECT [PRODUCT_ID]
      ,[WO_IDT]
      ,[SERIAL]
      ,[TERM_ID]
      ,[PASS_COUNT]
      ,[EVENT_DATE_TIME]
      ,[END_DATE]
      ,[START_DATE]
      ,[STATUS]
      ,[ORDER_ID]
      ,[LOT_ID]
      ,[LOT_NAME]
      ,[RowUpdated]
  FROM [SITMesDB].[dbo].[ARCH_T_SitMesComponentRT1A8997AF-5067-47d5-80DB-AF14C4BD2402/EV_$35$]
    WITH (nolock)
WHERE [EVENT_DATE_TIME] <= Convert (datetime,left(convert(varchar, getdate(), 23),20) + ' 06:00') AND [EVENT_DATE_TIME] >= Convert (datetime,left(convert(varchar, getdate()-1, 23),20) + ' 06:00') AND [TERM_ID] LIKE '%021%' 
"""

cursor.execute(script)

columns = [desc[0] for desc in cursor.description]
data = cursor.fetchall()
df = pd.read_sql_query(script, cnxn)
timestr = time.strftime("%d-%m-%Y_%H-%M-OK_NOK_ROTOR")
file_name = timestr
writer = pd.ExcelWriter(file_name+'.xlsx')
df.to_excel(writer, sheet_name='Pomiary', header = True, index = False)
writer.save() 
