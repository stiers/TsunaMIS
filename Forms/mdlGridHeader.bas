Attribute VB_Name = "Grid_Header"
'*******************************************************
'* TsunaMIS Grid Headers
'* Author: Ephramar A. Telog
'* Created: February 24, 2014
'* Email: ephramar@outlook.com
'*
'* Copyright 2014
'*******************************************************

Public Function GeneralLedgerHead()
    With IndexAccounting.grdGeneralLedger
        .Cols = 7: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "Date": .ColWidth(1) = 2200
        .Col = 2: .Text = "Line": .ColWidth(2) = 2200
        .Col = 3: .Text = "Description": .ColWidth(3) = 3200
        .Col = 4: .Text = "Amount": .ColWidth(4) = 2200
        .Col = 5: .Text = "Status": .ColWidth(5) = 2200
        .Col = 6: .Text = "Created By": .ColWidth(6) = 2200
    End With
End Function

Public Function SalesInvoiceHead()
    With SalesInvoice.grdSalesInvoice
        .Cols = 8: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "Date Created": .ColWidth(1) = 2200
        .Col = 2: .Text = "Date Due": .ColWidth(2) = 2200
        .Col = 3: .Text = "Quote Number": .ColWidth(3) = 2200
        .Col = 4: .Text = "Invoice Number": .ColWidth(4) = 2200
        .Col = 5: .Text = "Equipment": .ColWidth(5) = 2200
        .Col = 6: .Text = "Client": .ColWidth(6) = 2200
        .Col = 7: .Text = "Price": .ColWidth(7) = 2200
    End With
End Function

Public Function DailyActivityRecordHead()
    With IndexEngineering.grdDailyActivityRecord
        .Cols = 8: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "TSFR Number": .ColWidth(1) = 2000
        .Col = 2: .Text = "Client (Account)": .ColWidth(2) = 2200
        .Col = 3: .Text = "Equipment": .ColWidth(3) = 3200
        .Col = 4: .Text = "Job Type": .ColWidth(4) = 2200
        .Col = 5: .Text = "Contact Person": .ColWidth(5) = 3200
        .Col = 6: .Text = "Position": .ColWidth(6) = 2200
        .Col = 7: .Text = "Contact Number": .ColWidth(7) = 2200
    End With
End Function

Public Function PurchaseRequestHead()
    With IndexLogistics.grdDailyActivityRecord
        .Cols = 13: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "Date Issued": .ColWidth(1) = 2200
        .Col = 2: .Text = "Delivery Date": .ColWidth(2) = 2200
        .Col = 3: .Text = "Reference Number": .ColWidth(3) = 2200
        .Col = 4: .Text = "Invoice Number": .ColWidth(4) = 2200
        .Col = 5: .Text = "Client Name": .ColWidth(5) = 2200
        .Col = 6: .Text = "Client Address": .ColWidth(6) = 2200
        .Col = 7: .Text = "Telephone": .ColWidth(7) = 2200
        .Col = 8: .Text = "Cellphone": .ColWidth(8) = 2200
        .Col = 9: .Text = "Equipment": .ColWidth(9) = 2200
        .Col = 10: .Text = "Brand": .ColWidth(10) = 2200
        .Col = 11: .Text = "Model": .ColWidth(11) = 2200
        .Col = 12: .Text = "Product Number": .ColWidth(12) = 2200
    End With
End Function

Public Function SalesQuotationHead()
    With IndexSales.grdSalesQuote
        .Cols = 6: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "Quotation Number": .ColWidth(1) = 2000
        .Col = 2: .Text = "Date": .ColWidth(2) = 2200
        .Col = 3: .Text = "Equipment": .ColWidth(3) = 3200
        .Col = 4: .Text = "Client": .ColWidth(4) = 3200
        .Col = 5: .Text = "Total Amount": .ColWidth(5) = 2200
    End With
End Function

Public Function ServiceQuotationHead()
    With IndexSales.grdServiceQuote
        .Cols = 6: .Rows = 1
        .ColWidth(0) = 0
        
        .Col = 1: .Text = "Quotation Number": .ColWidth(1) = 2000
        .Col = 2: .Text = "Date": .ColWidth(2) = 2200
        .Col = 3: .Text = "Equipment": .ColWidth(3) = 3200
        .Col = 4: .Text = "Client": .ColWidth(4) = 3200
        .Col = 5: .Text = "Total Amount": .ColWidth(5) = 2200
    End With
End Function



