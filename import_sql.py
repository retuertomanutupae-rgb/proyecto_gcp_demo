import pandas
import pyodbc


#banco=pandas.read_excel("C:\\Users\\RichardRetuerto\\Desktop\\CLASES\\ZEGEL IPAE\Metodos\department.xlsx")
#print(banco.columns.values)

#server = 'yourservername' 
#database = 'AdventureWorks' 
#username = 'username' 
#password = 'yourpassword' 
#cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
#cursor = cnxn.cursor()


#cnxn = pyodbc.connect(r'Driver=SQL Server;Server=LAPTOP-4V2UNF33\SQLEXPRESS;Database=user_metodos;Trusted_Connection=yes;')
cnxn = pyodbc.connect(r'Driver=SQL Server;Server=10.210.12.34;Database=DB_DM_DASHBOARDLEALTAD;UID=usrDLealtad;PWD=Lealtad20;Trusted_Connection=no;')
cursor = cnxn.cursor()
#cursor.execute("Delete from DepartmentTest")
#cursor.execute("SELECT LastName FROM myContacts")
#while 1:
#    row = cursor.fetchone()
#    if not row:
#        break
#    print(row.LastName)
#cnxn.close()

# Insert Dataframe into SQL Server:
#for index, row in banco.iterrows():
#     cursor.execute("INSERT INTO DepartmentTest (DepartmentID,Name,GroupName) values(?,?,?)", row.DepartmentID, row.Name, row.GroupName)
#     print("Datos cargados a SQL SERVER")
#cnxn.commit()
#cursor.close()


cursor.execute("SELECT * FROM DepartmentTest")
while 1:
    row = cursor.fetchone()
    if not row:
       break
    print(row.DepartmentID, row.Name, row.GroupName)
cnxn.close()