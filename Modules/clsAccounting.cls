VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TsunaAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Terminate()
    Set Record = Nothing
    Set Connect = Nothing
End Sub

Public Function AddRecord(Line As String, Description As String, DateCreated As String, Amount As Double, Status As Integer, Invoice As String, Accountant As Integer)
    query = "INSERT INTO tbl_finance(line, description, date, amount, status, invoice, user_id) " & _
            "VALUES ('" & Line & "','" & Description & "','" & Format(DateCreated, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & Amount & "','" & Status & "','" & Invoice & "'," & Accountant & ")"
            
    Connect.Execute (query)
    
    MsgBox "Successfully added!", vbInformation, App.EXEName
End Function

Public Function UpdateRecord(Line As String, Description As String, DateCreated As String, Amount As Double, Status As Integer, Invoice As String, Accountant As Integer)
    query = "UPDATE `tbl_finance` SET `finance_id`=[value-1],`line`=[value-2],`description`=[value-3],`date`=[value-4],`amount`=[value-5],`status`=[value-6],`invoice`=[value-7],`currency`=[value-8],`user_id`=[value-9] WHERE 1"
End Function

Public Function DeleteRecord()
    query = "DELETE FROM `tbl_finance` WHERE 1"
End Function

Public Function DisplayToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT date, line, description, amount, status, display_name FROM tbl_finance " & _
            "JOIN tbl_users ON ID = user_id ORDER BY date DESC"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!Date: .CellAlignment = 1
            .Col = 2: .Text = Record!Line.value
            .Col = 3: .Text = Record!Description.value
            .Col = 4: .Text = Format(Record!Amount, "0.00"): .CellAlignment = 1
            .Col = 5: .Text = IIf(Record!Status = 1, "Income", "Expense")
            .Col = 6: .Text = Record!display_name
        End With
        Record.MoveNext
    Next counter
End Function

Public Function TotalAmount(Control As Label, Status As String)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT SUM(amount) AS total FROM tbl_finance WHERE status = '" & Status & "'"
    
    Record.Open query, Connect
    
    Control.Caption = Format(Record!total, "0.00")
End Function
