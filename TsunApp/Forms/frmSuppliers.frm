VERSION 5.00
Begin VB.Form frmSuppliers 
   Caption         =   "Suppliers"
   ClientHeight    =   8085
   ClientLeft      =   2250
   ClientTop       =   1260
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9660
   Begin VB.TextBox txtAddNewSupplier 
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CommandButton btnAddSupplier 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.ListBox lstSuppliers 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6330
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1440
      Width           =   4575
   End
   Begin VB.ComboBox cboBulkAction 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "frmSuppliers.frx":0000
      Left            =   240
      List            =   "frmSuppliers.frx":000D
      TabIndex        =   3
      Text            =   "Bulk Action"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton btnApplyAction 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton btnAddNewSupplier 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblAddNewSupplier 
      AutoSize        =   -1  'True
      Caption         =   "Add New Expense"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label lblSuppliers 
      AutoSize        =   -1  'True
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1170
   End
End
Attribute VB_Name = "frmSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddNewSupplier_Click()
    lblAddNewSupplier.Visible = True
    txtAddNewSupplier.Visible = True
    btnAddSupplier.Visible = True
End Sub

Private Sub btnAddSupplier_Click()
    Call Suppliers.AddNew(Trim(txtAddNewSupplier.Text))
    
    lblAddNewSupplier.Visible = False
    txtAddNewSupplier.Visible = False
    btnAddSupplier.Visible = False
    
    Call Suppliers.Display(lstSuppliers)
End Sub

Private Sub Form_Load()
    Call Suppliers.Display(lstSuppliers)
    
    lblAddNewSupplier.Visible = False
    txtAddNewSupplier.Visible = False
    btnAddSupplier.Visible = False
End Sub
