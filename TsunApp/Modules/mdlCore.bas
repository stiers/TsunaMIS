Attribute VB_Name = "mdlCore"
Public Connect As ADODB.Connection
Public Records As ADODB.Recordset

Public Users As clsUsers
Public ExpenseType As clsExpType
Public Suppliers As clsSuppliers
Public SalesLine As clsSalesLine
Public Ledger As clsLedger
Public MD5 As clsMD5

Public query As String

Sub Main()
    Set Connect = New ADODB.Connection
    Set Records = New ADODB.Recordset
    
    Set Users = New clsUsers
    Set ExpenseType = New clsExpType
    Set Suppliers = New clsSuppliers
    Set SalesLine = New clsSalesLine
    Set Ledger = New clsLedger
    Set MD5 = New clsMD5
    
    With Records
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With

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
