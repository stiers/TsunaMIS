VERSION 5.00
Begin VB.Form frmIndex 
   Caption         =   "Index"
   ClientHeight    =   8085
   ClientLeft      =   4785
   ClientTop       =   2310
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   7215
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
      Begin VB.Menu mnuSales 
         Caption         =   "Sales Lines"
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
    frmInputExp.Show
End Sub

Private Sub mnuInputSales_Click()
    frmInputSales.Show
End Sub

Private Sub mnuSales_Click()
    frmSalesLine.Show
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
