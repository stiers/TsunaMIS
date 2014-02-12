VERSION 5.00
Begin VB.Form frmAccExpType 
   Caption         =   "Add New Expense"
   ClientHeight    =   8085
   ClientLeft      =   4650
   ClientTop       =   1470
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   Begin VB.TextBox txtExpenseType 
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
      Left            =   3360
      TabIndex        =   1
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton btnExpenseType 
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
      Left            =   4560
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2310
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Expense Type"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1200
   End
End
Attribute VB_Name = "frmAccExpType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExpenseType_Click()
    query = "INSERT INTO tbl_tp_accounting_meta(meta_id, meta_type) VALUES ('','" & Trim(Me.txtExpenseType.Text) & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmAccExpType.Caption
    
    Me.txtExpenseType.Text = vbNullString
End Sub
