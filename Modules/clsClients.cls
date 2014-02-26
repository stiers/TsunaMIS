VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TsunaClients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Terminate()
    Set Record = Nothing
    Set Connect = Nothing
End Sub

Public Function AddClient(Name As String, Position As String, Company As String, Address As String, Email As String, Telephone As Long, Facsimile As Long, Mobile As Long, AdditionalInfo As String)

    query = "INSERT INTO tbl_tp_clients(client_name, client_position, client_company, client_address, client_email, client_telephone, client_fax, client_mobile, client_info) " & _
            "VALUES ('" & Trim(Name) & "','" & Trim(Position) & "','" & Trim(Company) & "','" & Trim(Address) & "','" & Trim(Email) & "','" & Telephone & "','" & Facsimile & "','" & Mobile & "','" & Trim(AdditionalInfo) & "')"
    
    Connect.Execute query
    
    MsgBox "Successfully added", vbInformation, "Message"
End Function

Public Function DisplayClientNameAsList(Control As ListBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT client_name FROM tbl_tp_clients"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!client_name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayCompanyNameAsCombo(Control As ComboBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT client_company FROM tbl_tp_clients"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!client_company
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function
