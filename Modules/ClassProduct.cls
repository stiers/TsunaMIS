VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function AddItem(EquipmentName As String, ProductName As String, FreightCost As String)
    If ProductName = "Choose nothing" Then ProductName = vbNullString
    
    query = "INSERT INTO tbl_tp_products(ID, equipment_name, product_line, freight_cost) VALUES ('','" & EquipmentName & "','" & ProductName & "','" & FreightCost & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmLogisticsPLAdd.Caption
End Function

Public Function DisplayDropdown(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT equipment_name FROM tbl_tp_products WHERE product_line = ''"
    
    Records.Open query, Connect
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        If Not Records.EOF Then
            For ComboData = 0 To Records.RecordCount - 1
                Combo.AddItem Records!equipment_name
                Records.MoveNext
            Next ComboData
        End If
    End If
End Function

Public Function DisplayGrid(Grid As MSFlexGrid)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT * FROM tbl_tp_products WHERE product_line <> ''"
    
    Records.Open query, Connect
    
    With Grid
        .Rows = Records.RecordCount
        .Cols = Records.Fields.Count
        
        .TextMatrix(0, 1) = "Equipment"
        .TextMatrix(0, 2) = "Product Line"
        .TextMatrix(0, 3) = "Freight Cost"
        
        For RowData = 1 To Records.RecordCount - 1
            For ColData = 0 To Records.Fields.Count - 1
                .TextMatrix(RowData, ColData) = Records.Fields(ColData)
            Next ColData
            Records.MoveNext
        Next RowData
        
        Dim x
        x = .Width / 3
        
        .ColWidth(0) = 350
        .ColWidth(1) = x
        .ColWidth(2) = x
        .ColWidth(3) = x
    End With
End Function

Public Function DisplayEquipment(Combo As ComboBox)
    If Records.State = 1 Then Records.Close
    
    query = "SELECT equipment_name FROM tbl_tp_products WHERE product_line <> ''"
    
    Records.Open query, Connect
    
    If Records.RecordCount > 0 Then
        Records.MoveFirst
        If Not Records.EOF Then
            For ComboData = 0 To Records.RecordCount - 1
                Combo.AddItem Records!equipment_name
                Records.MoveNext
            Next ComboData
        End If
    End If
End Function
