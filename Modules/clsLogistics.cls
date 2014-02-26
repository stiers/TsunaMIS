VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TsunaLogistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub Class_Terminate()
    Set Record = Nothing
    Set Connect = Nothing
End Sub

Public Function AddPurchase(DateIssued As Date, DateDelivery As Date, PurchaseNum As String, InvoiceNum As String, ClientName As String, ClientAddress As String, ClientNum As String, ClientMobile As String, EqName As String, EqBrand As String, EqModel As String, EqProductNum As String)
    query = "INSERT INTO tbl_tp_logistics(date_issued, date_delivery, reference_number, invoice_number, client_name, client_address, client_contact, client_mobile, eq_name, eq_brand, eq_model, eq_product_number) " & _
            "VALUES ('" & Format(DateIssued, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & Format(DateDelivery, "yyyy" & "-" & "MM" & "-" & "dd") & "','" & PurchaseNum & "','" & InvoiceNum & "','" & ClientName & "','" & ClientAddress & "','" & ClientNum & "','" & ClientMobile & "','" & EqName & "','" & EqBrand & "','" & EqModel & "','" & EqProductNum & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully added!", vbInformation, App.EXEName
End Function

Public Function DisplayBrandToCombo(Control As ComboBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT Item_Name FROM tbl_upcp_items"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!Item_Name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayProductToList(Control As ListBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT Item_Name FROM tbl_ucpc_items"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!Item_Name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayProductToCombo(Control As ComboBox)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT Item_Name FROM tbl_upcp_items"
    
    Record.Open query, Connect
    
    Control.Clear
    
    If Record.RecordCount > 0 Then
        Record.MoveFirst
        If Not Record.EOF Then
            For ControlData = 0 To Record.RecordCount - 1
                Control.AddItem Record!Item_Name
                Record.MoveNext
            Next ControlData
        End If
    End If
End Function

Public Function DisplayPurchaseRequests(Control As MSFlexGrid)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT date_issued, date_delivery, reference_number, invoice_number, client_name, client_address, client_contact, client_mobile, eq_name, eq_brand, eq_model, eq_product_number FROM tbl_tp_logistics"
    
    Record.Open query, Connect
    
    Dim counter As Integer
    
    For counter = 1 To Record.RecordCount
        With Control
            .Rows = .Rows + 1
            .Row = counter
            
            .Col = 1: .Text = Record!date_issued
            .Col = 2: .Text = Record!date_delivery
            .Col = 3: .Text = Record!reference_number
            .Col = 4: .Text = Record!invoice_number
            .Col = 5: .Text = Record!client_name
            .Col = 6: .Text = Record!client_address
            .Col = 7: .Text = Record!client_contact
            .Col = 8: .Text = Record!client_mobile
            .Col = 9: .Text = Record!eq_name
            .Col = 10: .Text = Record!eq_brand
            .Col = 11: .Text = Record!eq_model
            .Col = 12: .Text = Record!eq_product_number
        End With
        Record.MoveNext
    Next counter
End Function

Public Function DisplayToPurchaseRequestForm(Name As String)
    If Record.State = 1 Then Record.Close
    
    query = "SELECT client_mobile, client_address, client_telephone FROM tbl_tp_clients WHERE client_company = '" & Name & "'"
    
    Record.Open query, Connect
    
    With PurchaseRequest
        .Text_ReqClientMobile.Text = Record!client_mobile
        .Text_ReqClientDeliveryAdd = Record!client_address
        .Text_ReqClientNumber.Text = Record!client_telephone
    End With
End Function
