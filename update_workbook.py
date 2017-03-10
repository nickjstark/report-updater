#-------------------------------------------------------------------------------
# Purpose:      refresh Tableau packaged workbook file's underlying dataextract file from SQL
# Note:       could also load in from csv, Excel, etc.
#-------------------------------------------------------------------------------

import pyodbc, os, shutil, datetime, zipfile, time, sys
import dataextract as tde

t0 = time.time()

#helper dictionary for tde API data types
tdeTypes = {
    "<type 'int'>": 7, #integer
    "<type 'float'>": 10, #double
    "<type 'bool'>": 11, #boolean
    "<type 'datetime.date'>": 12, #date
    "<type 'datetime.datetime'>":13, #datetime
    "<type 'str'>" : 15, #char string
    "<type 'unicode'>":16 #unicode string
}

def add_row(r_ob, DType,col_i,v): #row, data type, col index, value
    if isinstance(v,int):
        r_ob.setInteger(col_i,v)
    elif isinstance(v,float):
        r_ob.setDouble(col_i,v)
    elif isinstance(v,datetime.datetime):
        r_ob.setDateTime(col_i,v.year, v.month, v.day, v.hour, v.minute, v.second, 0) #zero is milliseconds
    elif isinstance(v,datetime.date):
        r_ob.setDate(col_i,v.year, v.month, v.day)
    elif isinstance(v,unicode):
        r_ob.setString(col_i,v)
    elif isinstance(v,str):
        r_ob.setCharString(col_i,v)
    else:
        print "data type in add row function issue! " + str(type(v))

#Change variables below to math correct path, filenames, etc.
path = 'C:/YOUR/PATH'
dashboard_file = 'YOUR_DASH.twbx'
temp_ext = 'tableau_files' #folder to unzip to
tdeName = 'YOUR_EXTRACT.tde'
sqlQuery = "SELECT * FROM [DB].[dbo].[TABLE] WHERE [COL1] in ('vala', 'valb')" #just an ex.
updated_dashboard = 'UPDATED_WB.twbx' #updated twbx file

#pull data from warehouse
print "pull from SQL Server"
driver = '{SQL Server}'
server = 'YOUR_SERVER'
trusted_conn = 'yes'

conn = pyodbc.connect('driver=%s;server=%s;Trusted_Connection=%s'%(driver,server,trusted_conn))
cursor = conn.cursor()
cursor.execute(sqlQuery)

#unzip Packaged Workbook file to new folder
print "unzip workbook"
zip = zipfile.ZipFile(path + dashboard_file)
zip.extractall(path + temp_ext)
zip.close()
#remove existing extract file
os.remove(path + temp_ext + '/Data/Datasources/' + tdeName) #always in /Data/Datasources/extract_name.tde

#create new tableau extract w/ orignal name
print "create tableau extract"
os.chdir(path + temp_ext + '/Data/Datasources/')
tdefile = tde.Extract(tdeName) #create tde file
tableDef = tde.TableDefinition() #create a new table def

for column in cursor.description:
    tableDef.addColumn(column[0],tdeTypes[str(column[1])])

tdetable = tdefile.addTable("Extract",tableDef) #has to be Extract for some reason

print "adding rows"
for row in cursor.fetchall():
    i = 0
    newrow = tde.Row(tableDef)
    while i < len(row):
        val = row[i]
        t = type(row[i])
        if val == 'None' or val is None:
            newrow.setNull(i)
        else:
            add_row(newrow,t,i,val)
        i += 1
    tdetable.insert(newrow) #insert row into TDE table
    newrow.close()

#close SQL connection and extract file
tdefile.close()
conn.close()

#rezip all files and folders, maintains folder structure in zipped workbook
def zip_dir(zipname, dir_to_zip):
    dir_to_zip_len = len(dir_to_zip.rstrip(os.sep)) + 1
    with zipfile.ZipFile(zipname, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
        for dirname, subdirs, files in os.walk(dir_to_zip):
            for filename in files:
                path = os.path.join(dirname, filename)
                entry = path[dir_to_zip_len:]
                zf.write(path, entry)

print "repackaging Tableau workbook"
zip_dir(path + updated_dashboard, path + temp_ext)
os.chdir(path)
shutil.rmtree(path + temp_ext)
print round((time.time() - t0)/60,2), ' min process time'