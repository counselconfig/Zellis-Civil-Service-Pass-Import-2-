print("LOADING OPTIMUM CONVERSION & CARDS IMPORT...\n")

import win32com.client as win32
from getch import pause #https://stackoverflow.com/questions/11876618/python-press-any-key-to-exit
from getch import pause_exit
import os
from sqlalchemy import create_engine
import urllib
import pandas as pd
import datetime as dt
import pyodbc #needed for the exe build even though not used explicitly below
import logging
from datetime import datetime




#https://stackoverflow.com/questions/9918646/how-to-convert-xls-to-xlsx
#fname = r"C:\Temp\cardsdataimport_python\xlstoxlsx\Security3.xls"

#print("OPTIMUM CONVERSION & CARDS IMPORT\n")
print("1.) Please close all Excel programs")
print("2.) There must be only one xls file in the directory")



pause(message="3.) Press any key to convert")


# Create a logging instance https://stackoverflow.com/questions/55169364/python-how-to-write-error-in-the-console-in-txt-file
logger = logging.getLogger("my_application")
logger.setLevel(logging.INFO) # you can set this to be DEBUG, INFO, ERROR


# Assign a file-handler to that instance
fh = logging.FileHandler(datetime.now().strftime("mylogfile_%H_%M_%d_%m_%Y.log"), delay=True) # https://stackoverflow.com/questions/4180518/make-a-python-log-file-only-when-there-are-errors-using-logging-module
fh.setLevel(logging.INFO) # again, you can set this differently

# Format your logs (optional)
formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
fh.setFormatter(formatter) # This will set the format to the file handler

# Add the handler to your logging instance
logger.addHandler(fh)


try:

    # Get a single xls file by searching for the extension alone
    the_dir = os.getcwd()
    all_xls_files = list(filter(lambda x: x.endswith('.xls'), os.listdir(the_dir))) # https://stackoverflow.com/questions/3964681/find-all-files-in-a-directory-with-extension-txt-in-python
    fn = (", ".join(all_xls_files)) # https://stackoverflow.com/questions/13207697/how-to-remove-square-brackets-from-list-in-python
    fname = os.path.abspath(fn)
    
    # Get name of file being used to print # https://openwritings.net/pg/python/python-how-get-filename-without-extension
    path=fname # Get the filename only from the initial file path.
    filename = os.path.basename(path) # Use splitext() to get filename and extension separately.
    (file, ext) = os.path.splitext(filename)

    #fname = os.path.abspath("optimum.xls") # https://appdividend.com/2021/06/07/python-relative-path/#:~:text=A%20relative%20path%20that%20depicts,provide%20a%20full%20absolute%20path.
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    print("\nConversion to " + file + ".xlsx complete")
    
except Exception as e: # https://stackoverflow.com/questions/3383865/how-to-log-error-to-file-and-not-fail-on-exception
    logger.exception(e) # Will send the errors to the file
    pause_exit(status=0, message="\nError finding the file, see log and raise a ticket. Press any key to exit")
    



print("Importing data to the Cards database...")

# Settings
#TargetServer = "na-t-sqlc02v01\sql1"
TargetServer = "na-sqlc03v01,4334"
SchemaName = "dbo"
TargetDb = "Cards"
TableName = "Badge"
#TableName = "test" #do not add dbo. otherwise will try to create a table
UserName = "EstatesAccess" #needs db_datawriter membership privileges in MSSQL
Password = "Eur0paL3ague"
#SourceFile = "C:\\Temp\\excelimport\\test.xlsx"
SourceFile = os.path.abspath(file + ".xlsx")

# Configure the Connection
Params = urllib.parse.quote_plus(r"DRIVER={SQL Server};SERVER=" + TargetServer + ";DATABASE=" + TargetDb + ";UID=" + UserName + ";PWD=" + Password)
ConnStr = "mssql+pyodbc:///?odbc_connect={}".format(Params)
Engine = create_engine(ConnStr)

# Load the sheet into a DataFrame
#df = pd.read_excel(SourceFile, sheet_name = "Sheet1", header = 0)
df = pd.read_excel(SourceFile, header = 3, dtype=str) # read from 4th row down   na_filter=False https://stackoverflow.com/questions/62325180/how-i-can-stop-python-read-excel-to-convert-date-to-datetime-kindly-help-post
#df['Badge Expiry'] = pd.to_datetime(df['Badge Expiry']).dt.date




# Clear the Data in Target Table
sql = "Truncate Table Badge"
with Engine.begin() as conn:
    conn.execute(sql)

try:
    # Load the Data in DataFrame into Table
    df.to_sql(TableName, con=Engine, schema=SchemaName, if_exists="append", index=False)

    print(dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S") + " | Data imported successfully")
    pause_exit(status=0, message="\n4.) Press any key to exit")

except Exception as e:
    logger.exception(e)
    pause_exit(status=0, message="\nData import error, see log and raise a ticket. Press any key to exit")



#https://stackoverflow.com/questions/21174956/insert-permission-was-denied-on-the-object-employee-info-database-payroll-s