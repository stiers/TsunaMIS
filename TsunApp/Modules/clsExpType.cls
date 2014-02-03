VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Display(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT post_title FROM tbl_posts WHERE post_type = 'wpa_expense_type'"
    
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Title"
        
        For RowData = 1 To Records.RecordCount
            .TextMatrix(RowData, 0) = RowData
            
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData + 1) = Records.Fields(ColData)
            Next ColData
            
            Records.MoveNext
            
        Next RowData
        
        x = .Width / 4
        
        .ColWidth(0) = 350
        .ColWidth(1) = x
        
    End With
End Function

Public Function Dropdown(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT post_title FROM tbl_posts WHERE post_type = 'wpa_expense_type'"
    
    Records.Open query, Connect
    
    Combo.Clear
    
    For ctr = 0 To Records.RecordCount
    
        If Not Records.EOF Then
            Combo.AddItem Records!post_title
            Records.MoveNext
        End If
        
    Next ctr
End Function
