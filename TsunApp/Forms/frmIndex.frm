VERSION 5.00
Begin VB.Form frmIndex 
   BackColor       =   &H80000005&
   Caption         =   "Index"
   ClientHeight    =   8685
   ClientLeft      =   3225
   ClientTop       =   1710
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   9855
   WindowState     =   2  'Maximized
   Begin VB.PictureBox imgLogo 
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   9480
      Picture         =   "frmIndex.frx":0000
      ScaleHeight     =   641
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   641
      TabIndex        =   0
      Top             =   960
      Width           =   9615
   End
   Begin VB.Menu mnuAccounting 
      Caption         =   "Accounting"
      Begin VB.Menu mnuAccEntries 
         Caption         =   "Accounting Entries"
         Begin VB.Menu mnuInputSales 
            Caption         =   "Input Sales"
         End
         Begin VB.Menu mnuInputExp 
            Caption         =   "Input Expenses"
         End
      End
      Begin VB.Menu mnuIncome 
         Caption         =   "Income Statement"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "Ledger"
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
      End
      Begin VB.Menu mnuExpType 
         Caption         =   "Expense Types"
      End
      Begin VB.Menu mnuSalesLine 
         Caption         =   "Sales Lines"
      End
   End
   Begin VB.Menu mnuQuotation 
      Caption         =   "Quotations"
      Begin VB.Menu mnuViewQuotations 
         Caption         =   "Display All Quotations"
      End
      Begin VB.Menu mnuSalesQuote 
         Caption         =   "Sales Quotation"
      End
      Begin VB.Menu mnuSvcQuotePart 
         Caption         =   "Service Quotation (Parts)"
      End
      Begin VB.Menu mnuSvcQuotePMS 
         Caption         =   "Service Quotation (Maintenance)"
      End
      Begin VB.Menu mnuSalesReport 
         Caption         =   "Sales Report"
      End
      Begin VB.Menu mnuSvcReport 
         Caption         =   "Service Report"
      End
   End
   Begin VB.Menu mnuPersonnel 
      Caption         =   "Personnel"
      Begin VB.Menu mnuDTR 
         Caption         =   "Daily Time Record"
      End
      Begin VB.Menu mnuPayroll 
         Caption         =   "Payroll"
      End
   End
   Begin VB.Menu mnuLogistics 
      Caption         =   "Logistics"
      Begin VB.Menu mnuPurchaseReq 
         Caption         =   "Purchase Request"
      End
   End
   Begin VB.Menu mnuUsers 
      Caption         =   "Users"
      Begin VB.Menu mnuUserAll 
         Caption         =   "All Users"
      End
      Begin VB.Menu mnuUserNew 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnuUserProfile 
         Caption         =   "Your Profile"
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuExpType_Click()
    frmExpType.Show
End Sub

Private Sub mnuInputExp_Click()
    frmInputExp.Show vbModal, Me
End Sub

Private Sub mnuInputSales_Click()
    frmInputSales.Show vbModal, Me
End Sub

Private Sub mnuLedger_Click()
    frmLedger.Show vbModal, Me
End Sub

Private Sub mnuSalesLine_Click()
    frmSalesLine.Show
End Sub

Private Sub mnuSalesQuote_Click()
    frmSalesQuote.Show
End Sub

Private Sub mnuSuppliers_Click()
    frmSuppliers.Show
End Sub

Private Sub mnuUserAll_Click()
    frmUsers.Show
End Sub

Private Sub mnuUserNew_Click()
    frmUserNew.Show
End Sub

Private Sub mnuUserProfile_Click()
    frmProfile.Show
End Sub

Private Sub mnuViewQuotations_Click()
    frmQuotations.Show
End Sub
