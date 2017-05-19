import xlrd
import os
import pyodbc
import datetime
import sched


#establish MS DB Connection
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=XXX;DATABASE=XXX;UID=XXX;PWD=XXX')
#create Cursor
cursor = cnxn.cursor()
#Drop table statement for trunc/reload
cursor.execute("IF OBJECT_ID('dbo.Package_Rate_Change', 'U') IS NOT NULL DROP TABLE XXX.dbo.Package_Rate_Change;")
#create table
cursor.execute('''CREATE TABLE Package_Rate_Change
              (pol_num varchar(20),
               product varchar(50),
               insd_name varchar(500),
               renewal_date datetime,
               expiring_policy_premium varchar(50),
               renewal_policy_premium varchar(50),
               expiring_property_premium varchar(50),
               renewal_property_premium varchar(50),
               expiring_liability_premium varchar(50),
               renewal_liability_premium varchar(50),
               property_rpc varchar(50),
               liability_rpc varchar(50),
               total_rpc varchar(50),
               date_inserted datetime
               )''')

file_location = "S:\Commercial Lines\Rate Monitor\itd_files"
query = """INSERT INTO Package_Rate_Change (pol_num,
                                            product,
                                            insd_name,
                                            renewal_date,
                                            expiring_policy_premium,
                                            renewal_policy_premium,
                                            expiring_property_premium,
                                            renewal_property_premium,
                                            expiring_liability_premium,
                                            renewal_liability_premium,
                                            property_rpc,
                                            liability_rpc,
                                            total_rpc,
                                            date_inserted
                                            )
                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?);"""


def handleUnicode(i):
    if type(i) == unicode:
        return str(i)
    elif type(i) == float:
        return int(i)
    else:
        return i

for i in os.listdir(file_location):
    book = xlrd.open_workbook(os.path.join(file_location, i))
    sheet = book.sheet_by_index(0)
    pol_num = handleUnicode(sheet.cell_value(1, 6))
    product = sheet.cell_value(3, 6).replace("-"," - ")
    insd_name = sheet.cell_value(1, 1)
    renewal_date = sheet.cell_value(3, 1)
    time_tuple = datetime.datetime(*xlrd.xldate_as_tuple(renewal_date, book.datemode))
    expiring_policy_premium = sheet.cell_value(5, 1)
    renewal_policy_premium = sheet.cell_value(6, 1)
    expiring_property_premium = sheet.cell_value(13, 1)
    renewal_property_premium = sheet.cell_value(13, 2)
    expiring_liability_premium = sheet.cell_value(13, 6)
    renewal_liability_premium = sheet.cell_value(13, 7)
    property_rpc = sheet.cell_value(22, 2)
    liability_rpc = sheet.cell_value(28, 7)
    total_rpc = sheet.cell_value(22, 2)
    date_inserted = datetime.datetime.utcnow()

    values = (handleUnicode(pol_num), product, insd_name, time_tuple, expiring_policy_premium,
              renewal_policy_premium, expiring_property_premium, renewal_property_premium,
              expiring_liability_premium, renewal_liability_premium,
              property_rpc, liability_rpc, total_rpc, date_inserted)
    cursor.execute(query, values)
    print values


cursor.close()
cnxn.commit()
cnxn.close()
