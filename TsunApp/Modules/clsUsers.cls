VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Display(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close

    query = "SELECT user_login, display_name, user_email, meta_value " & _
            "FROM tbl_users " & _
            "RIGHT JOIN tbl_usermeta ON ID = user_id " & _
            "WHERE meta_key = 'tbl_user_level'"

    Records.Open query, Connect

    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1

        .TextMatrix(0, 1) = "Username"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 3) = "E-mail"
        .TextMatrix(0, 4) = "Role"

        For RowData = 1 To Records.RecordCount
            .TextMatrix(RowData, 0) = RowData
            
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData + 1) = Records.Fields(ColData)
            Next ColData
            
            Records.MoveNext
        Next RowData
        
        X = .Width / 4
        
        .ColWidth(0) = 350
        .ColWidth(1) = X
        .ColWidth(2) = X
        .ColWidth(3) = X
        .ColWidth(4) = 1900
    End With
End Function

Public Function Dropdown(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT display_name FROM tbl_users"
    
    Records.Open query, Connect
    
    Combo.Clear
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        Do Until Records.EOF Or Records.BOF
            With Combo
                .AddItem Records!display_name
            End With
        Records.MoveNext
        Loop
    End If
End Function
