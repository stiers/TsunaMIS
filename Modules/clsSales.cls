VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TsunaSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Terminate()
    Set Record = Nothing
    Set Connect = Nothing
End Sub

Public Function AddInvoice(DateCreated As Date, DateDue As Date, RefNum As String, InvNum As String, Equipment As String, Client As String, Price As Double)
    query = "INSERT INTO tbl_tp_invoice(si_date_created, si_date_due, si_ref_num, si_number, si_equipment, si_client, si_price) VALUES ('" & Format(DateCreated, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & Format(DateDue, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & RefNum & "','" & InvNum & "','" & Equipment & "','" & Client & "','" & Price & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully added!", vbInformation, App.EXEName
End Function

Public Function AddSales(Number As Integer, DateCreated As Date, SalesType As String, Equipment As String, Client As String, NetPrice As Double, Freight As Double, Tax As Double, GrossProfit As Double, OpEx As Double, Representation As Double, Training As Double, Miscellaneous As Double, Delivery As Double, OtherCharge As Double, TotalAmount As Double)
    query = "INSERT INTO tbl_tp_quotations(q_ref_num, q_date, q_type, q_equipment, q_client, q_price, q_freight, q_tax, q_profit, q_opex, q_rep, q_training, q_misc, q_delivery, q_other, total_amount) " & _
            "VALUES ('" & Number & "','" & Format(DateCreated, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & SalesType & "','" & Equipment & "','" & Client & "','" & NetPrice & "','" & Freight & "','" & Tax & "','" & GrossProfit & "','" & OpEx & "','" & Representation & "','" & Training & "','" & Miscellaneous & "','" & Delivery & "','" & OtherCharge & "','" & TotalAmount & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully added!", vbInformation, App.EXEName
End Function

Public Function UpdateSales(Number As Integer, DateCreated As Date, SalesType As String, Equipment As String, Client As String, NetPrice As Double, Freight As Double, Tax As Double, GrossProfit As Double, OpEx As Double, Representation As Double, Training As Double, Miscellaneous As Double, Delivery As Double, OtherCharge As Double, TotalAmount As Double)
    query = "UPDATE tbl_tp_quotations SET q_ref_num='" & Number & "',q_date='" & Format(DateCreated, "yyyy" & "-" & "MM" & "-" & "dd") & "',q_type='" & SalesType & "',q_equipment='" & Equipment & "',q_client='" & Client & "',q_price='" & NetPrice & "',q_freight='" & Freight & "',q_tax='" & Tax & "',q_profit='" & GrossProfit & "',q_opex='" & OpEx & "',q_rep='" & Representation & "',q_training='" & Training & "',q_misc='" & Miscellaneous & "',q_delivery='" & Delivery & "',q_other='" & OtherCharge & "',total_amount='" & TotalAmount & "' " & _
            "WHERE ID='" & CurrentQuotationID & "'"
    
    Connect.Execute (query)
    
    MsgBox "Successfully updated!", vbInformation, App.EXEName
End Function

Public Function DeleteSales()
    'query "DELETE FROM tbl_tp_quotations WHERE ID = '" & CurrentQuotationID & "'"
End Function

Public Function DisplaySalesQuotesToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT q_ref_num, q_date, q_equipment, q_client, total_amount FROM tbl_tp_quotations WHERE q_type = 'sales'"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!q_ref_num: .CellAlignment = 1
            .Col = 2: .Text = Record!q_date: .CellAlignment = 1
            .Col = 3: .Text = Record!q_equipment
            .Col = 4: .Text = Record!q_client
            .Col = 5: .Text = Record!total_amount: .CellAlignment = 1
        End With
        Record.MoveNext
    Next counter
End Function

Public Function DisplayServiceQuotesToGrid(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT q_ref_num, q_date, q_equipment, q_client, total_amount FROM tbl_tp_quotations WHERE q_type = 'maintenance'"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!q_ref_num: .CellAlignment = 1
            .Col = 2: .Text = Record!q_date: .CellAlignment = 1
            .Col = 3: .Text = Record!q_equipment
            .Col = 4: .Text = Record!q_client
            .Col = 5: .Text = Record!total_amount: .CellAlignment = 1
        End With
        Record.MoveNext
    Next counter
End Function

Public Function DisplayInvoice(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT si_date_created, si_date_due, si_ref_num, si_number, si_equipment, si_client, si_price FROM tbl_tp_invoice"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!si_date_created
            .Col = 2: .Text = Record!si_date_due
            .Col = 3: .Text = Record!si_ref_num
            .Col = 4: .Text = Record!si_number
            .Col = 5: .Text = Record!si_equipment
            .Col = 6: .Text = Record!si_client
            .Col = 7: .Text = Record!si_price
        End With
        Record.MoveNext
    Next counter
End Function

Public Function ToDouble(TextBox As TextBox)
    If TextBox = vbNullString Then TextBox.Text = Format(TextBox.Text, "0.00")
End Function
