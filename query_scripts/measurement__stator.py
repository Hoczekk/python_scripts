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
SELECT
      [EQUIPMENT_ID]
      ,[DEVICE]
      ,[TIMESTAMP]
      ,[ORDER_ID]
      ,[SERIAL_NUMBER]
      ,[TERMINAL_ID]
      ,[WO_ID]
      ,[MEASUREMENT_ID]
      ,[MEASUREMENT_TYPE]
      ,[UPPER_LIMIT]
      ,[MEASUREMENT]
      ,[LOWER_LIMIT]
      ,[TARGET_VALUE]
      ,[UNIT]
      ,[STATUS]
      ,[MACHINE_CYCLE]
  FROM [SITMesDB].[dbo].[ARCH_T_SitMesComponentRT1A8997AF-5067-47d5-80DB-AF14C4BD2402/EC_MEASUREMENTS_$86$]
    WITH (nolock)
WHERE [TIMESTAMP] <= Convert (datetime,left(convert(varchar, getdate(), 23),20) + ' 06:00') AND [TIMESTAMP] >= Convert (datetime,left(convert(varchar, getdate()-1, 23),20) + ' 06:00') AND [TERMINAL_ID] LIKE 'EC031%' 
"""

cursor.execute(script)

columns = [desc[0] for desc in cursor.description]
data = cursor.fetchall()
df = pd.read_sql_query(script, cnxn)
timestr = time.strftime("%d-%m-%Y_%H-%M-POMIARY_STATOR")
file_name = timestr
writer = pd.ExcelWriter(file_name+'.xlsx')
df.to_excel(writer, sheet_name='Pomiary', header = True, index = False)
writer.save() 
