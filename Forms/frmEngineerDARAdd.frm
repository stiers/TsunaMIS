VERSION 5.00
Begin VB.Form frmEngineerDARAdd 
   Caption         =   "Add Record"
   ClientHeight    =   8085
   ClientLeft      =   5760
   ClientTop       =   1965
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
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
      Left            =   3360
      TabIndex        =   16
      Top             =   2280
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
      ItemData        =   "frmEngineerDARAdd.frx":0000
      Left            =   3360
      List            =   "frmEngineerDARAdd.frx":000D
      TabIndex        =   14
      Top             =   2760
      Width           =   2775
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   3360
      TabIndex        =   11
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtTSFRNum 
         Height          =   405
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
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
      Left            =   4560
      TabIndex        =   10
      Top             =   6480
      Width           =   1575
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
      Left            =   3360
      TabIndex        =   3
      Top             =   5640
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
      Left            =   3360
      TabIndex        =   2
      Top             =   5160
      Width           =   2775
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
      Left            =   3360
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   3105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   1440
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   690
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmEngineerDARAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddDAR_Click()
    query = "INSERT INTO tbl_tp_service(ID, TSFR, job_account, job_equipment, job_type, contact_person, contact_position, contact_number) VALUES ('', '" & Me.txtTSFRNum & "', '" & Me.cboTSFRAcc.Text & "', '" & Me.cboTSFREq.Text & "', '" & Me.cboTSFRJob.Text & "', '" & Me.txtTSFRName.Text & "', '" & Me.txtTSFRPosition.Text & "', '" & Me.txtTSFRContact.Text & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmEngineerDAR.Caption
End Sub

Private Sub Form_Load()
    Call Client.DisplayDropdown(cboTSFRAcc)
    Call Product.DisplayEquipment(cboTSFREq)
End Sub
