Attribute VB_Name = "Core"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Connect As ADODB.Connection
Public Records As ADODB.Recordset

Public MD5 As ClassMD5

Public Accounting As ClassAccounting
Public Client As ClassClient
Public User As ClassUser
Public Product As ClassProduct
Public Quotation As ClassQuotes

Public Tax As Double
Public Profit As Double
Public OpEx As Double
Public Rep As Double

Sub Main()
    Set Connect = New ADODB.Connection
    Set Records = New ADODB.Recordset
    
    Set MD5 = New ClassMD5
    
    Set Accounting = New ClassAccounting
    Set Client = New ClassClient
    Set User = New ClassUser
    Set Product = New ClassProduct
    Set Quotation = New ClassQuotes
    
    Tax = 0.15
    Profit = 0.31
    OpEx = 0.14
    Rep = 0.02
    
    With Records
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    
    ConnectionString = "Driver={Mysql ODBC 3.51 Driver};" & _
                       "Server=localhost;" & _
                       "Port=3306;" & _
                       "Database=tsuna;" & _
                       "User=root;" & _
                       "Password=2n329fdx;" & _
                       "Option=3;"
    
    Connect.Open ConnectionString
    
    Tsuna_Initialize
End Sub

Public Function Tsuna_Initialize() As Boolean
    Dim Username As String, Password As String
    Dim Fail As Boolean, Successful As Boolean
    
    Randomize
    
    Username = GetSetting(App.EXEName, "Settings", "LastUser", "")
    
    Fail = frmSecurity.GetLogIn(Username, Password, frmIndex)
    
    Do While Fail
        If Records.State = 1 Then Records.Close
        
        query = "SELECT ID, user_login, user_pass FROM tbl_users WHERE user_login = '" & Replace(Username, "'", "''") & "'"
        
        Records.Open query, Connect
        
        If Records.RecordCount = 0 Then GoTo Bye
        
        If LCase(MD5.DigestStrToHexStr(Password)) = Records!user_pass Then
            LogInUserID = Records!ID
            LogInUserName = Records!user_login
            
            SaveSetting App.EXEName, "Setting", "LastUser", Records!user_login
            
            Successful = True
            frmIndex.Show
            Exit Do
        End If
        
Bye:
        If Not Successful Then
            Fail = False
            
            If MsgBox("Invalid login, do you want to try again ?", vbQuestion + vbYesNo, "Invalid Login") = vbYes Then
                Sleep 200 + 300 * Rnd
                Fail = frmSecurity.GetLogIn(Username, Password, frmIndex)
            End If
        End If
    Loop
    
    Tsuna_Initialize = Successful
End Function
