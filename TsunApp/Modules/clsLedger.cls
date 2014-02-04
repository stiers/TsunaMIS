VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Display(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT  ledger_date, post_title, amount " & _
            "FROM tbl_wpaccounting_ledger " & _
            "RIGHT JOIN tbl_posts ON type_id = ID " & _
            "WHERE ledger_date IS NOT NULL"
            
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Type"
        .TextMatrix(0, 3) = "Amount"
        
        For RowData = 1 To Records.RecordCount
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
    End With
End Function
