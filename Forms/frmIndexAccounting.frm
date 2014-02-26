VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IndexAccounting 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Accounting"
   ClientHeight    =   9450
   ClientLeft      =   1470
   ClientTop       =   1125
   ClientWidth     =   16440
   Icon            =   "frmIndexAccounting.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo_AccountType 
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
      ItemData        =   "frmIndexAccounting.frx":058A
      Left            =   5520
      List            =   "frmIndexAccounting.frx":0594
      TabIndex        =   30
      Top             =   8280
      Width           =   2775
   End
   Begin VB.CommandButton CommandButton_Print 
      BackColor       =   &H80000005&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CommandButton_Add 
      BackColor       =   &H8000000D&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CommandButton_Update 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton CommandButton_Delete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text_AccountAmount 
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
      Left            =   5520
      TabIndex        =   11
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text_AccountTitle 
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
      Left            =   5520
      TabIndex        =   10
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text_AccountInvNum 
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
      Left            =   5520
      TabIndex        =   4
      Top             =   6600
      Width           =   2775
   End
   Begin VB.TextBox Text_AccountDescription 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   7200
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid grdGeneralLedger 
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   5530
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker_AccountCreated 
      Height          =   375
      Left            =   5520
      TabIndex        =   13
      Top             =   4800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   108134401
      CurrentDate     =   41675
   End
   Begin VB.Label Label_SystemDate 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label_SystemDate"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13440
      TabIndex        =   29
      Top             =   9000
      Width           =   1365
   End
   Begin VB.Label Label_SystemTime 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label_SystemTime"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14880
      TabIndex        =   28
      Top             =   9000
      Width           =   1350
   End
   Begin VB.Label Label_AccountTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label Label_AccountExpense 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      TabIndex        =   26
      Top             =   7320
      Width           =   360
   End
   Begin VB.Label Label_AccountIncome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      TabIndex        =   25
      Top             =   6240
      Width           =   360
   End
   Begin VB.Label Label_AccountBalance 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      TabIndex        =   24
      Top             =   5160
      Width           =   360
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance from previous period"
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
      TabIndex        =   23
      Top             =   4800
      Width           =   2595
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Income"
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
      TabIndex        =   22
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense"
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
      TabIndex        =   21
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overview"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   990
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   18
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   8160
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
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
      Left            =   3720
      TabIndex        =   9
      Top             =   6600
      Width           =   1380
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   7200
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   3720
      TabIndex        =   7
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Left            =   3720
      TabIndex        =   6
      Top             =   5400
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Created"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "General Ledger"
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
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome user! You have no pending requests at this moment"
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
      Left            =   10560
      TabIndex        =   1
      Top             =   240
      Width           =   5400
   End
   Begin VB.Menu mnuLoanMgt 
      Caption         =   "Loan Management"
      Begin VB.Menu mnuLoanMgtAdd 
         Caption         =   "Add New Loan"
      End
   End
   Begin VB.Menu mnuSalesInvoice 
      Caption         =   "Sales Invoice"
   End
End
Attribute VB_Name = "IndexAccounting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Add_Click()
    Call Accounting.AddRecord(Me.Text_AccountTitle.Text, Me.Text_AccountDescription.Text, Me.DTPicker_AccountCreated.value, Me.Text_AccountAmount.Text, IIf(Me.Combo_AccountType.Text = "Income", 1, 2), Me.Text_AccountInvNum.Text, CurrentUserID)
    Call GeneralLedgerHead
    Call Accounting.DisplayToGrid(grdGeneralLedger)
End Sub

Private Sub Form_Activate()
    Call Accounting.TotalAmount(Label_AccountIncome, 1)
    Call Accounting.TotalAmount(Label_AccountExpense, 2)
    
    Me.Label_AccountTotal.Caption = Format((CDbl(Me.Label_AccountIncome.Caption) - CDbl(Me.Label_AccountExpense.Caption)), "0.00")
End Sub

Private Sub Form_Click()
    Me.Label9.Caption = "Add New"
    
    Me.CommandButton_Add.Visible = True
    Me.CommandButton_Update.Visible = False
    Me.CommandButton_Delete.Visible = False
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.Label_SystemDate.Caption = Format$(Now, "m/d/yy")
    Me.Label_SystemTime.Caption = Format$(Now, "hh:mm AM/PM")
    
    'hide update and delete buttons
    Me.CommandButton_Update.Visible = False
    Me.CommandButton_Delete.Visible = False
    
    'from Grid Header module
    Call GeneralLedgerHead
    Call Accounting.DisplayToGrid(grdGeneralLedger)
End Sub

Private Sub grdGeneralLedger_DblClick()
    Me.CommandButton_Add.Visible = False
    Me.CommandButton_Update.Visible = True
    Me.CommandButton_Delete.Visible = True
    
    Me.Label9.Caption = "Update"
    
    If Record.State = 1 Then Record.Close

    query = "SELECT finance_id FROM tbl_finance WHERE date = '" & Format(Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 1), "yyyy" & "-" & "MM" & "-" & "dd") & "'"

    Record.Open query, Connect

    CurrentFinanceID = Record!finance_id
    
    Me.DTPicker_AccountCreated.value = Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 1)
    Me.Text_AccountTitle.Text = Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 2)
    Me.Text_AccountDescription.Text = Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 3)
    Me.Text_AccountAmount.Text = Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 4)
    Me.Combo_AccountType.Text = Me.grdGeneralLedger.TextMatrix(Me.grdGeneralLedger.Row, 5)
    Me.Text_AccountInvNum.Text = Record!finance_id
End Sub

Private Sub mnuSalesInvoice_Click()
    SalesInvoice.Show
End Sub
