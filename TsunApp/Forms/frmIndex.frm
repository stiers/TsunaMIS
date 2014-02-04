VERSION 5.00
Begin VB.Form frmIndex 
   BackColor       =   &H80000005&
   Caption         =   "Index"
   ClientHeight    =   8685
   ClientLeft      =   3015
   ClientTop       =   1275
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   19200
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
      Begin VB.Menu mnuSales 
         Caption         =   "Sales Lines"
      End
   End
   Begin VB.Menu mnuLogistics 
      Caption         =   "Logistics"
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

Private Sub mnuPersonnel_Click()

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
