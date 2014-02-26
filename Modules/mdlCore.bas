Attribute VB_Name = "Core"
'*******************************************************
'* TsunaMIS Core Version 3.0
'* Author: Ephramar A. Telog
'* Created: February 6, 2014
'* Email: ephramar@outlook.com
'*
'* Copyright 2014
'*******************************************************

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'ADODB Variables
Public Connect As ADODB.Connection
Public Record As ADODB.Recordset

'SQL String
Public query As String

'Global User Variables
Public CurrentUserID As Integer
Public CurrentUserName As String
Public CurrentUserGroup As Integer

'Global Accounting Variable
Public CurrentFinanceID As Integer

'Global Sales Variable
Public CurrentQuotationID As Integer

'Global Engineering Variable
Public CurrentRAID As Integer

'Global Logistics Variable
Public pID As Integer

'Class Variables
Public Accounting As TsunaAccounting
Public Client As TsunaClients
Public Engineering As TsunaEngineering
Public Logistics As TsunaLogistics
Public MD5 As TsunaMD5
Public Sales As TsunaSales

Sub Main()
    Dim ConnectionString As String
    
    Set Connect = New ADODB.Connection
    Set Record = New ADODB.Recordset
    
    Set Accounting = New TsunaAccounting
    Set Client = New TsunaClients
    Set Engineering = New TsunaEngineering
    Set Logistics = New TsunaLogistics
    Set MD5 = New TsunaMD5
    Set Sales = New TsunaSales
    
    With Record
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With
    
    ConnectionString = "Driver={Mysql ODBC 3.51 Driver};" & _
                       "Server=localhost;" & _
                       "Port=3306;" & _
                       "Database=db_tsuna;" & _
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
    
    Fail = Security.GetUserInfo(Username, Password, Index)
    
    Do While Fail
        If Record.State = 1 Then Record.Close
        
        query = "SELECT tbl_users.ID, user_login, user_pass, group_id FROM tbl_users INNER JOIN tbl_bp_groups_members ON user_id = tbl_users.ID WHERE user_login = '" & Replace(Username, "'", "''") & "'"
        
        Record.Open query, Connect
        
        If Record.RecordCount = 0 Then GoTo Bye
        
        If LCase(MD5.DigestStrToHexStr(Password)) = Record!user_pass Then
            CurrentUserID = Record!ID
            CurrentUserName = Record!user_login
            CurrentUserGroup = Record!group_id
            
            SaveSetting App.EXEName, "Setting", "LastUser", CurrentUserName
            
            Successful = True
            
            If CurrentUserID > 1 Then
                If CurrentUserGroup = 1 Then
                    IndexAccounting.Show
                ElseIf CurrentUserGroup = 2 Then
                    IndexEngineering.Show
                ElseIf CurrentUserGroup = 3 Then
                    IndexSales.Show
                Else
                    IndexLogistics.Show
                End If
            Else
                Index.Show
            End If
            
            Exit Do
        End If
        
Bye:
        If Not Successful Then
            Fail = False
            
            If MsgBox("Invalid username or password" & vbCrLf & "Do you want to try again?", vbQuestion + vbYesNo, Security.Caption) = vbYes Then
                Sleep 200 + 300 * Rnd
                Fail = Security.GetUserInfo(Username, Password, Index)
            End If
        End If
    Loop
    
    Tsuna_Initialize = Successful
End Function

