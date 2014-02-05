VERSION 5.00
Begin VB.Form frmInputSales 
   Caption         =   "Input Sales"
   ClientHeight    =   8055
   ClientLeft      =   3645
   ClientTop       =   1365
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   6585
   Begin VB.CommandButton btnInputSales 
      Caption         =   "Input Sales"
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
      Left            =   240
      TabIndex        =   15
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   14
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   12
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   11
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   10
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtUserUserName 
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
      TabIndex        =   8
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales7 
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
      ItemData        =   "frmInputSales.frx":0000
      Left            =   240
      List            =   "frmInputSales.frx":0002
      TabIndex        =   7
      Text            =   "Choose type"
      Top             =   5280
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales6 
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
      ItemData        =   "frmInputSales.frx":0004
      Left            =   240
      List            =   "frmInputSales.frx":0006
      TabIndex        =   6
      Text            =   "Choose type"
      Top             =   4560
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales5 
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
      ItemData        =   "frmInputSales.frx":0008
      Left            =   240
      List            =   "frmInputSales.frx":000A
      TabIndex        =   5
      Text            =   "Choose type"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales4 
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
      ItemData        =   "frmInputSales.frx":000C
      Left            =   240
      List            =   "frmInputSales.frx":000E
      TabIndex        =   4
      Text            =   "Choose type"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales3 
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
      ItemData        =   "frmInputSales.frx":0010
      Left            =   240
      List            =   "frmInputSales.frx":0012
      TabIndex        =   3
      Text            =   "Choose type"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales2 
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
      ItemData        =   "frmInputSales.frx":0014
      Left            =   240
      List            =   "frmInputSales.frx":0016
      TabIndex        =   2
      Text            =   "Choose type"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.ComboBox cboInputSales1 
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
      ItemData        =   "frmInputSales.frx":0018
      Left            =   240
      List            =   "frmInputSales.frx":001A
      TabIndex        =   0
      Text            =   "Choose type"
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblInputSales 
      AutoSize        =   -1  'True
      Caption         =   "Input Sales"
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
      Width           =   1410
   End
End
Attribute VB_Name = "frmInputSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SalesLine.Dropdown(cboInputSales1)
    Call SalesLine.Dropdown(cboInputSales2)
    Call SalesLine.Dropdown(cboInputSales3)
    Call SalesLine.Dropdown(cboInputSales4)
    Call SalesLine.Dropdown(cboInputSales5)
    Call SalesLine.Dropdown(cboInputSales6)
    Call SalesLine.Dropdown(cboInputSales7)
End Sub
