VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayList(List As ListBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT client_name FROM tbl_tp_clients"
    
    Records.Open query, Connect
    
    List.Clear
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        Do Until Records.EOF
            With List
                .AddItem Records!client_name
            End With
        Records.MoveNext
        Loop
    End If
End Function

Public Function DisplayDropdown(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT client_name FROM tbl_tp_clients"
    
    Records.Open query, Connect
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        If Not Records.EOF Then
            For ComboData = 0 To Records.RecordCount - 1
                Combo.AddItem Records!client_name
                Records.MoveNext
            Next ComboData
        End If
    End If
End Function