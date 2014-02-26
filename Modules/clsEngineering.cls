VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TsunaEngineering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Terminate()
    Set Record = Nothing
    Set Connect = Nothing
End Sub

Public Function AddRecord(ReferenceNumber As String, Account As String, Equipment As String, JobOrder As String, ContactName As String, ContactPosition As String, ContactNumber As String)
    query = "INSERT INTO tbl_tp_service(TSFR, job_account, job_equipment, job_type, contact_person, contact_position, contact_number) " & _
            "VALUES ('" & Trim(ReferenceNumber) & "','" & Account & "','" & Equipment & "','" & JobOrder & "','" & Trim(ContactName) & "','" & Trim(ContactPosition) & "','" & Trim(ContactNumber) & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully added!", vbInformation, App.EXEName
End Function

Public Function UpdateRecord(ReferenceNumber As String, Account As String, Equipment As String, JobOrder As String, ContactName As String, ContactPosition As String, ContactNumber As String)
    query = "UPDATE tbl_tp_service SET TSFR='" & Trim(ReferenceNumber) & "',job_account='" & Account & "',job_equipment='" & Equipment & "',job_type='" & JobOrder & "',contact_person='" & Trim(ContactName) & "',contact_position='" & Trim(ContactPosition) & "',contact_number='" & Trim(ContactNumber) & "' WHERE ID = '" & CurrentRAID & "'"
    
    Connect.Execute (query)
    
    MsgBox "Successfully updated!", vbInformation, App.EXEName
End Function

Public Function DeleteRecord()
    query = "DELETE FROM tbl_tp_service WHERE ID = '" & CurrentRAID & "'"
    
    Connect.Execute (query)
    
    MsgBox "Successfully deleted!", vbInformation, App.EXEName
End Function

Public Function DisplayToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT TSFR, job_account, job_equipment, job_type, contact_person, contact_position, contact_number FROM tbl_tp_service"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!TSFR: .CellAlignment = 1
            .Col = 2: .Text = Record!job_account
            .Col = 3: .Text = Record!job_equipment
            .Col = 4: .Text = Record!job_type
            .Col = 5: .Text = Record!contact_person
            .Col = 6: .Text = Record!contact_position
            .Col = 7: .Text = Record!contact_number: .CellAlignment = 1
        End With
        Record.MoveNext
    Next counter
End Function

