VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IndexEngineering 
   BackColor       =   &H80000005&
   Caption         =   "Engineering"
   ClientHeight    =   8280
   ClientLeft      =   1455
   ClientTop       =   2010
   ClientWidth     =   12735
   Icon            =   "frmIndexEngineering.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12735
   Begin VB.CommandButton Command1 
      Caption         =   "Add Report"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
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
      Left            =   9000
      TabIndex        =   4
      Text            =   "Search "
      Top             =   4320
      Width           =   3495
   End
   Begin VB.CommandButton btnServiceQuotationAdd 
      Caption         =   "Add Report"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Alignment       =   1  'Right Justify
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
      Left            =   9000
      TabIndex        =   0
      Text            =   "Search "
      Top             =   360
      Width           =   3495
   End
   Begin MSFlexGridLib.MSFlexGrid grdDailyActivityRecord 
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   4260
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   240
      TabIndex        =   6
      Top             =   4920
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   4471
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
      Left            =   9600
      TabIndex        =   9
      Top             =   7800
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
      Left            =   11040
      TabIndex        =   8
      Top             =   7800
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipment History"
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
      TabIndex        =   7
      Top             =   4200
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Daily Activity Record"
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
      Width           =   2700
   End
End
Attribute VB_Name = "IndexEngineering"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnServiceQuotationAdd_Click()
    EngineeringDAR.Show
End Sub

Private Sub Form_Activate()
    'from Grid Header module
    Call DailyActivityRecordHead

    Call Engineering.DisplayToGrid(grdDailyActivityRecord)
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.Label_SystemDate.Caption = Format$(Now, "m/d/yy")
    Me.Label_SystemTime.Caption = Format$(Now, "hh:mm AM/PM")
End Sub

Private Sub grdDailyActivityRecord_DblClick()
    EngineeringDAR.Show
    
    With EngineeringDAR
        If Record.State = 1 Then Record.Close
        
        query = "SELECT * FROM tbl_tp_service WHERE TSFR = '" & grdDailyActivityRecord.TextMatrix(grdDailyActivityRecord.Row, 1) & "'"
        
        Record.Open query, Connect
        
        .Label1.Caption = "Edit Record"
        
        CurrentRAID = Record!ID     'check the Core module for usage
        .txtTSFRNum.Text = Record!TSFR
        .cboTSFRAcc.Text = Record!job_account
        .cboTSFREq.Text = Record!job_equipment
        .cboTSFRJob.Text = Record!job_type
        .txtTSFRName.Text = Record!contact_person
        .txtTSFRPosition.Text = Record!contact_position
        .txtTSFRContact.Text = Record!contact_number
    End With
End Sub
