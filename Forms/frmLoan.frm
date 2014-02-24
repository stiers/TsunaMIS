VERSION 5.00
Begin VB.Form frmLoan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Loan"
   ClientHeight    =   7185
   ClientLeft      =   2220
   ClientTop       =   855
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   8610
   Begin VB.TextBox txtLoanPeriod 
      Height          =   495
      Left            =   3120
      TabIndex        =   20
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtLoanInterest 
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox txtLoanTotal 
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   5640
      Width           =   5055
   End
   Begin VB.TextBox txtLoanRepayAmt 
      Height          =   495
      Left            =   3120
      TabIndex        =   16
      Top             =   4800
      Width           =   5055
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6720
      TabIndex        =   15
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000005&
      Caption         =   "Period"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   3840
      Width           =   3015
      Begin VB.OptionButton optMonths 
         BackColor       =   &H80000005&
         Caption         =   "Months"
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
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optYears 
         BackColor       =   &H80000005&
         Caption         =   "Years"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox txtLoanAmount 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2280
      Width           =   5055
   End
   Begin VB.ComboBox cboDepartment 
      Height          =   435
      Left            =   3120
      TabIndex        =   1
      Top             =   1560
      Width           =   5055
   End
   Begin VB.ComboBox cboUsers 
      Height          =   435
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loan"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department:"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount of Loan:"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Charged:"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Period of Loan:"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   3960
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Repayment:"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   9
      Left            =   240
      TabIndex        =   8
      Top             =   4800
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Repayment:"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   240
      TabIndex        =   7
      Top             =   5640
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   675
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalculate_Click()
    Dim Repay As Double
    Dim Repayment As Double
    
    If txtLoanAmount.Text = "" Then
        MsgBox "Enter Loan Amount", vbInformation
        Exit Sub
    End If
    
    If txtLoanInterest.Text = "" Then
        MsgBox "Set Interest", vbInformation
        Exit Sub
    End If
    
    If txtLoanPeriod.Text = "" Then
        MsgBox "Enter Loan Period", vbInformation
        Exit Sub
    End If

    Repay = (txtLoanAmount.Text * txtLoanInterest.Text * txtLoanPeriod.Text) / 100
    Repayment = Format(Repay + txtLoanAmount, "##,###.00")
    
    If optYears.value = True Then
        txtLoanRepayAmt.Text = (Repayment) / (txtLoanPeriod * 12)
    End If
    If optMonths.value = True Then
        txtLoanRepayAmt.Text = (Repayment / txtLoanPeriod)
    End If
    
    txtLoanTotal.Text = Repayment
End Sub

Private Sub btnSave_Click()
    query = "INSERT INTO tbl_finance(finance_id, line, description, date, amount, status, invoice, currency, user_id) " & _
            "VALUES ('','Employee Loan','Loan of " & Me.cboUsers.Text & "','" & Date & "','" & Me.txtLoanTotal & "','2','0034','','1')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmLoan.Caption
End Sub

Private Sub Form_Load()
    Call User.DisplayDropdown(cboUsers)
End Sub
