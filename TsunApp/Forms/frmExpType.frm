VERSION 5.00
Begin VB.Form frmExpType 
   Caption         =   "Expense Types"
   ClientHeight    =   8085
   ClientLeft      =   4470
   ClientTop       =   1440
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9690
   Begin VB.CommandButton btnAddExp 
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
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtAddNewExpType 
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   4215
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
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboBulkAction 
      Height          =   375
      ItemData        =   "frmExpType.frx":0000
      Left            =   240
      List            =   "frmExpType.frx":000D
      TabIndex        =   3
      Text            =   "Bulk Action"
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox lstExpType 
      Height          =   6330
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1440
      Width           =   4575
   End
   Begin VB.CommandButton btnAddNewExpType 
      Caption         =   "Add Expense"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblAddNewExpType 
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
      TabIndex        =   5
      Top             =   1440
      Width           =   2310
   End
   Begin VB.Label lblExpType 
      AutoSize        =   -1  'True
      Caption         =   "Expense Types"
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
      Width           =   1860
   End
End
Attribute VB_Name = "frmExpType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddExp_Click()
    Call ExpenseType.AddNew(Trim(txtAddNewExpType.Text))
    
    lblAddNewExpType.Visible = False
    txtAddNewExpType.Visible = False
    btnAddExp.Visible = False
    
    Call ExpenseType.Display(lstExpType)
End Sub

Private Sub btnAddNewExpType_Click()
    lblAddNewExpType.Visible = True
    txtAddNewExpType.Visible = True
    btnAddExp.Visible = True
End Sub

Private Sub Form_Load()
    Call ExpenseType.Display(lstExpType)
    
    lblAddNewExpType.Visible = False
    txtAddNewExpType.Visible = False
    btnAddExp.Visible = False
End Sub
