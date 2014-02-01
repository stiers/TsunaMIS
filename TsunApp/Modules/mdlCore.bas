Attribute VB_Name = "mdlCore"
Public Connect As ADODB.Connection
Public Records As ADODB.Recordset

Public query As String

Sub Main()
    Set Connect = New ADODB.Connection
    Set Records = New ADODB.Recordset

    Dim ConnectionString As String
    
    ConnectionString = "Driver={Mysql ODBC 3.51 Driver};" & _
                       "Server=localhost;" & _
                       "Port=3306;" & _
                       "Database=db_tsuna;" & _
                       "User=root;" & _
                       "Password=2n329fdx;" & _
                       "Option=3;"
    
    Connect.Open ConnectionString
    
    frmIndex.Show
End Sub
