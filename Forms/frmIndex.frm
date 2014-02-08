VERSION 5.00
Begin VB.Form frmIndex 
   BackColor       =   &H80000005&
   Caption         =   "Index"
   ClientHeight    =   8085
   ClientLeft      =   25335
   ClientTop       =   2640
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   15180
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAccounting 
      Caption         =   "Accounting"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuLedger 
         Caption         =   "Ledger"
      End
      Begin VB.Menu mnuIncome 
         Caption         =   "Income"
      End
      Begin VB.Menu mnuPayroll 
         Caption         =   "Payroll"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetAccounting 
         Caption         =   "Settings"
         Begin VB.Menu mnuAddExp 
            Caption         =   "Add New Expense"
         End
      End
   End
   Begin VB.Menu mnuSales 
      Caption         =   "Sales"
      Begin VB.Menu mnuQuotation 
         Caption         =   "Quotations"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
      End
      Begin VB.Menu mnuClients 
         Caption         =   "Clients"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetSales 
         Caption         =   "Settings"
         Begin VB.Menu mnuAddSalesQuote 
            Caption         =   "Add Sales Quotation"
         End
         Begin VB.Menu mnuAddSvcQuoteParts 
            Caption         =   "Add Service Quotation (Parts)"
         End
         Begin VB.Menu mnuAddSvcQuoteMaint 
            Caption         =   "Add Service Quotation (Maintenance)"
         End
         Begin VB.Menu mnuAddClient 
            Caption         =   "Add Client"
         End
      End
   End
   Begin VB.Menu mnuEngineering 
      Caption         =   "Engineering"
      Begin VB.Menu mnuDAR 
         Caption         =   "Daily Activity Record"
      End
      Begin VB.Menu mnuVEH 
         Caption         =   "View Equipment History"
      End
   End
   Begin VB.Menu mnuLogistics 
      Caption         =   "Logistics"
      Begin VB.Menu mnuPRQ 
         Caption         =   "Purchase Request"
      End
      Begin VB.Menu mnuProductLine 
         Caption         =   "Product Line"
         Begin VB.Menu mnuProductLineList 
            Caption         =   "List of Products"
         End
         Begin VB.Menu mnuProductLineAdd 
            Caption         =   "Add New Product"
         End
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      NegotiatePosition=   3  'Right
      WindowList      =   -1  'True
      Begin VB.Menu mnuUserGroups 
         Caption         =   "User Groups"
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAddClient_Click()
    frmSalesClientAdd.Show
End Sub

Private Sub mnuAddExp_Click()
    frmAccExpType.Show
End Sub

Private Sub mnuAddSalesQuote_Click()
    frmSalesQuotationAddSales.Show
End Sub

Private Sub mnuAddSvcQuoteMaint_Click()
    frmSalesQuotationAddSvcM.Show
End Sub

Private Sub mnuAddSvcQuoteParts_Click()
    frmSalesQuotationAddSvcP.Show
End Sub

Private Sub mnuClients_Click()
    frmSalesClient.Show
End Sub

Private Sub mnuDAR_Click()
    frmEngineerDAR.Show
End Sub

Private Sub mnuIncome_Click()
    frmAccIncome.Show
End Sub

Private Sub mnuLedger_Click()
    frmAccLedger.Show
End Sub

Private Sub mnuPayroll_Click()
    frmAccPayroll.Show
End Sub

Private Sub mnuProductLineAdd_Click()
    frmLogisticsPLAdd.Show
End Sub

Private Sub mnuProductLineList_Click()
    frmLogisticsPL.Show
End Sub

Private Sub mnuPRQ_Click()
    frmLogisticsPR.Show
End Sub

Private Sub mnuQuotation_Click()
    frmSalesQuotation.Show
End Sub

Private Sub mnuUserGroups_Click()
    frmUserGroups.Show
End Sub

Private Sub mnuVEH_Click()
    frmEngineerVEH.Show
End Sub
