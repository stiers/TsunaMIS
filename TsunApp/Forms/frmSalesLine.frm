VERSION 5.00
Begin VB.Form frmSalesLine 
   Caption         =   "Sales Line"
   ClientHeight    =   8085
   ClientLeft      =   3660
   ClientTop       =   2025
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9510
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
      Top             =   960
      Width           =   975
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
      ItemData        =   "frmSalesLine.frx":0000
      Left            =   240
      List            =   "frmSalesLine.frx":000D
      TabIndex        =   3
      Text            =   "Bulk Action"
      Top             =   960
      Width           =   1815
   End
   Begin VB.ListBox lstSalesLine 
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
      TabIndex        =   2
      Top             =   1560
      Width           =   4575
   End
   Begin VB.CommandButton btnAddNewSalesLine 
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
   Begin VB.Label lblSalesLine 
      AutoSize        =   -1  'True
      Caption         =   "Sales Line"
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
      Width           =   1245
   End
End
Attribute VB_Name = "frmSalesLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SalesLine.Display(lstSalesLine)
End Sub
