VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayGrid(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT line, description, date, amount, status, invoice, user_id FROM `tbl_finance`"
    
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Subject"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Date"
        .TextMatrix(0, 4) = "Amount"
        .TextMatrix(0, 5) = "Type"
        .TextMatrix(0, 6) = "Invoice"
        .TextMatrix(0, 7) = "Accountant"
        
        For RowData = 1 To Records.RecordCount
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData + 1) = Records.Fields(ColData)
            Next ColData
            Records.MoveNext
        Next RowData
        
        Dim x
        x = .width / 7
        
        .ColWidth(0) = 350
        .ColWidth(1) = x
        .ColWidth(2) = x
        .ColWidth(3) = x
        .ColWidth(4) = x
        .ColWidth(5) = x
        .ColWidth(6) = x
        .ColWidth(7) = x
    End With
End Function
