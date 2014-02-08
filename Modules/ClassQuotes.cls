VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function DisplayGrid(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT q_ref_num, q_date, q_equipment, q_client, q_price FROM tbl_tp_quotations"
    
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Number"
        .TextMatrix(0, 2) = "Date"
        .TextMatrix(0, 3) = "Equipment"
        .TextMatrix(0, 4) = "Client"
        .TextMatrix(0, 5) = "Price"
        
        For RowData = 1 To Records.RecordCount
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData + 1) = Records.Fields(ColData)
            Next ColData
            Records.MoveNext
        Next RowData
        
        Dim x
        x = .width / 6
        
        .ColWidth(0) = 350
        .ColWidth(1) = x
        .ColWidth(2) = x
        .ColWidth(3) = x
        .ColWidth(4) = x
        .ColWidth(5) = x
    End With
End Function

Public Function DisplayRequest(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT q_ref_num, q_date, q_equipment, q_client, client_address, total_amount FROM tbl_tp_quotations " & _
            "INNER JOIN tbl_tp_clients ON client_name = q_client"
    
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount + 1
        .Cols = Records.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Number"
        .TextMatrix(0, 2) = "Date"
        .TextMatrix(0, 3) = "Equipment"
        .TextMatrix(0, 4) = "Client"
        .TextMatrix(0, 5) = "Address"
        .TextMatrix(0, 6) = "Price"
        
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
    End With
End Function

Public Function ToDouble(TextBox As TextBox)
    If TextBox = vbNullString Then TextBox.Text = Format(0, "0.00")
End Function
