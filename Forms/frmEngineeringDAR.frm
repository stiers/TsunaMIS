VERSION 5.00
Begin VB.Form EngineeringDAR 
   BackColor       =   &H80000005&
   Caption         =   "Daily Activity Record"
   ClientHeight    =   8580
   ClientLeft      =   3615
   ClientTop       =   1590
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10110
   Begin VB.CommandButton btnDeleteDAR 
      Caption         =   "Delete DAR"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton btnUpdateDAR 
      Caption         =   "Update DAR"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox txtTSFRName 
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
      Left            =   4200
      TabIndex        =   8
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtTSFRPosition 
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
      Left            =   4200
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txtTSFRContact 
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
      Left            =   4200
      TabIndex        =   6
      Top             =   5880
      Width           =   2775
   End
   Begin VB.CommandButton btnAddDAR 
      Caption         =   "Add DAR"
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
      Left            =   480
      TabIndex        =   5
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TSFR Number"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   2775
      Begin VB.TextBox txtTSFRNum 
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.ComboBox cboTSFRAcc 
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
      Left            =   4200
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cboTSFRJob 
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
      ItemData        =   "frmEngineeringDAR.frx":0000
      Left            =   4200
      List            =   "frmEngineeringDAR.frx":0010
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
   Begin VB.ComboBox cboTSFREq 
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
      Left            =   4200
      TabIndex        =   0
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add Record"
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
      Left            =   480
      TabIndex        =   16
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   15
      Top             =   4920
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account"
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
      Left            =   480
      TabIndex        =   14
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Job Order"
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
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   1620
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Left            =   480
      TabIndex        =   12
      Top             =   5400
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   480
      TabIndex        =   11
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Person who signed TSFR"
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
      Left            =   480
      TabIndex        =   10
      Top             =   3960
      Width           =   3105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equipment"
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
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   930
   End
End
Attribute VB_Name = "EngineeringDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strAccountName As String

Private Sub btnAddDAR_Click()
    Call Engineering.AddRecord(Me.txtTSFRNum.Text, Me.cboTSFRAcc.Text, Me.cboTSFREq.Text, Me.cboTSFRJob.Text, Me.txtTSFRName.Text, Me.txtTSFRPosition.Text, Me.txtTSFRContact.Text)
    Unload Me
End Sub

Private Sub cboTSFRAcc_LostFocus()
    strAccountName = cboTSFRAcc.Text
    
    If Record.State = 1 Then Record.Close
    
    query = "SELECT contact_person, contact_position, contact_number FROM tbl_tp_service"
    
    Record.Open query, Connect
    
    Me.txtTSFRName.Text = Record!contact_person
    Me.txtTSFRPosition.Text = Record!contact_position
    Me.txtTSFRContact.Text = Record!contact_number
End Sub

Private Sub Form_Activate()
    If Not Me.Label1.Caption = "Add Record" Then
        Me.btnAddDAR.Visible = False
        Me.btnUpdateDAR.Visible = True
        Me.btnDeleteDAR.Visible = True
    Else
        Me.btnUpdateDAR.Visible = False
        Me.btnDeleteDAR.Visible = False
    End If
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.btnUpdateDAR.Visible = False
    Me.btnDeleteDAR.Visible = False
    
    'from TsunaClients class
    Call Client.DisplayCompanyNameAsCombo(cboTSFRAcc)
    
    'from TsunaLogistics class
    Call Logistics.DisplayProductToCombo(cboTSFREq)
End Sub
