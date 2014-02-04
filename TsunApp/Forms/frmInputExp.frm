VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInputExp 
   Caption         =   "Input Expenses"
   ClientHeight    =   7950
   ClientLeft      =   2460
   ClientTop       =   2040
   ClientWidth     =   15765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   15765
   Begin VB.CommandButton btnInputExp 
      Caption         =   "Input Expenses"
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
      TabIndex        =   36
      Top             =   6600
      Width           =   1575
   End
   Begin VB.ComboBox cboSupplier7 
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
      ItemData        =   "frmInputExp.frx":0000
      Left            =   6480
      List            =   "frmInputExp.frx":0002
      TabIndex        =   35
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote7 
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
      Left            =   12720
      TabIndex        =   33
      Top             =   5760
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp7 
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
      ItemData        =   "frmInputExp.frx":0004
      Left            =   240
      List            =   "frmInputExp.frx":0006
      TabIndex        =   32
      Text            =   "Materials"
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt7 
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
      TabIndex        =   31
      Top             =   5760
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier6 
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
      ItemData        =   "frmInputExp.frx":0008
      Left            =   6480
      List            =   "frmInputExp.frx":000A
      TabIndex        =   30
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote6 
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
      Left            =   12720
      TabIndex        =   28
      Top             =   5040
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp6 
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
      ItemData        =   "frmInputExp.frx":000C
      Left            =   240
      List            =   "frmInputExp.frx":000E
      TabIndex        =   27
      Text            =   "Marketing"
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt6 
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
      TabIndex        =   26
      Top             =   5040
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier5 
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
      ItemData        =   "frmInputExp.frx":0010
      Left            =   6480
      List            =   "frmInputExp.frx":0012
      TabIndex        =   25
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote5 
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
      Left            =   12720
      TabIndex        =   23
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp5 
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
      ItemData        =   "frmInputExp.frx":0014
      Left            =   240
      List            =   "frmInputExp.frx":0016
      TabIndex        =   22
      Text            =   "Loan"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt5 
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
      TabIndex        =   21
      Top             =   4320
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier4 
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
      ItemData        =   "frmInputExp.frx":0018
      Left            =   6480
      List            =   "frmInputExp.frx":001A
      TabIndex        =   20
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote4 
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
      Left            =   12720
      TabIndex        =   18
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp4 
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
      ItemData        =   "frmInputExp.frx":001C
      Left            =   240
      List            =   "frmInputExp.frx":001E
      TabIndex        =   17
      Text            =   "General"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt4 
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
      TabIndex        =   16
      Top             =   3600
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier3 
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
      ItemData        =   "frmInputExp.frx":0020
      Left            =   6480
      List            =   "frmInputExp.frx":0022
      TabIndex        =   15
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote3 
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
      Left            =   12720
      TabIndex        =   13
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp3 
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
      ItemData        =   "frmInputExp.frx":0024
      Left            =   240
      List            =   "frmInputExp.frx":0026
      TabIndex        =   12
      Text            =   "Fees"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt3 
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
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier2 
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
      ItemData        =   "frmInputExp.frx":0028
      Left            =   6480
      List            =   "frmInputExp.frx":002A
      TabIndex        =   10
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtExpNote2 
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
      Left            =   12720
      TabIndex        =   8
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp2 
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
      ItemData        =   "frmInputExp.frx":002C
      Left            =   240
      List            =   "frmInputExp.frx":002E
      TabIndex        =   7
      Text            =   "Equipment"
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt2 
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
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.ComboBox cboSupplier1 
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
      ItemData        =   "frmInputExp.frx":0030
      Left            =   6480
      List            =   "frmInputExp.frx":0032
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dteExpDate1 
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin VB.TextBox txtExpNote1 
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
      Left            =   12720
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cboInputExp1 
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
      ItemData        =   "frmInputExp.frx":0034
      Left            =   240
      List            =   "frmInputExp.frx":0036
      TabIndex        =   1
      Text            =   "Donation"
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtExpAmt1 
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
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dteExpDate2 
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin MSComCtl2.DTPicker dteExpDate3 
      Height          =   375
      Left            =   9600
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin MSComCtl2.DTPicker dteExpDate4 
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   3600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin MSComCtl2.DTPicker dteExpDate5 
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin MSComCtl2.DTPicker dteExpDate6 
      Height          =   375
      Left            =   9600
      TabIndex        =   29
      Top             =   5040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin MSComCtl2.DTPicker dteExpDate7 
      Height          =   375
      Left            =   9600
      TabIndex        =   34
      Top             =   5760
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      Format          =   7995393
      CurrentDate     =   41673
   End
   Begin VB.Label lblExpNotes 
      AutoSize        =   -1  'True
      Caption         =   "Note"
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
      Left            =   13920
      TabIndex        =   41
      Top             =   960
      Width           =   435
   End
   Begin VB.Label lblExpDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Left            =   10680
      TabIndex        =   40
      Top             =   960
      Width           =   405
   End
   Begin VB.Label lblSupplier 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
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
      Left            =   7440
      TabIndex        =   39
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblExpAmt 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
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
      Left            =   4440
      TabIndex        =   38
      Top             =   960
      Width           =   675
   End
   Begin VB.Label lblExpense 
      AutoSize        =   -1  'True
      Caption         =   "Expense"
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
      Left            =   1200
      TabIndex        =   37
      Top             =   960
      Width           =   720
   End
   Begin VB.Label lblInputExp 
      AutoSize        =   -1  'True
      Caption         =   "Input Expenses"
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
      Width           =   1935
   End
End
Attribute VB_Name = "frmInputExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call ExpenseType.Dropdown(cboInputExp1)
    Call ExpenseType.Dropdown(cboInputExp2)
    Call ExpenseType.Dropdown(cboInputExp3)
    Call ExpenseType.Dropdown(cboInputExp4)
    Call ExpenseType.Dropdown(cboInputExp5)
    Call ExpenseType.Dropdown(cboInputExp6)
    Call ExpenseType.Dropdown(cboInputExp7)

    Call Suppliers.Dropdown(cboSupplier1)
    Call Suppliers.Dropdown(cboSupplier2)
    Call Suppliers.Dropdown(cboSupplier3)
    Call Suppliers.Dropdown(cboSupplier4)
    Call Suppliers.Dropdown(cboSupplier5)
    Call Suppliers.Dropdown(cboSupplier6)
    Call Suppliers.Dropdown(cboSupplier7)
End Sub
