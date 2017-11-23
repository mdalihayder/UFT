
'******************************************************************************************************************************************************************************************************************************************************
																		'Data Base Connection
'******************************************************************************************************************************************************************************************************************************************************


'Option Explicit Force variable declaration
Option Explicit

Function FnDataBase(TestValue)
	
Dim Con, Rs, SQL, Record

'Create description object for connection
Set Con=CreateObject("ADODB.Connection")

'Create description object for recordset
Set Rs=CreateObject("ADODB.Recordset")

'Write the query
SQL = TestValue
Reporter.ReportEvent micDone, "SQL Query"," SQL Query executed, Query is: "& SQL

'Connect with the SQL Server Database with valid credentials
Con.Open "Provider=SQLOLEDB.1;Data Source=MIRCLSQASQL01;User ID=webuser;Password=3edfGMjjyY876rtgYTRE46J66K8jmYRRDXe4;Persist Security Info=True;Initial Catalog=JPay;"

'Run the query while connected to the database
Rs.Open SQL,Con

''Retrieve the data
'Do While Not Rs.EOF
'
'Record = Rs(TestValue2)
'
'Msgbox Record
'
'Rs.MoveNext
'
'Loop
'
'Release objects'Release objects
Set Rs    = Nothing
Set Con    = Nothing

End Function
'FnDataBase "Update InmateAccount Set status='2' where InmateID ='0000000100' and PermLoc ='S_50000222'"





'******************************************************************************************************************************************************************************************************************************************************
																		'End Data Base Connection
'******************************************************************************************************************************************************************************************************************************************************

