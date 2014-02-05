VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Display(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close

    query = "SELECT * FROM tbl_quotations"

    Records.Open query, Connect

    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count

        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Reference Number"
        .TextMatrix(0, 2) = "Kind of Quote"

        For RowData = 1 To Records.RecordCount
            .TextMatrix(RowData, 0) = Records!q_id
            
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData) = Records.Fields(ColData)
            Next ColData
            
            Records.MoveNext
        Next RowData
        
        X = .Width / 3
        
        .ColWidth(0) = 300
        .ColWidth(1) = X
        .ColWidth(2) = X
    End With
End Function
