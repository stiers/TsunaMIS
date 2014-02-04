VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   8085
   ClientLeft      =   2925
   ClientTop       =   1275
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9585
   Begin MSFlexGridLib.MSFlexGrid grdUsers 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   12938
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
   End
   Begin VB.CommandButton btnAddNewUser 
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
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblUsers 
      AutoSize        =   -1  'True
      Caption         =   "Users"
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
      Width           =   690
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddNewUser_Click()
    frmUserNew.Show
End Sub

Private Sub Form_Load()
    Call Users.Display(grdUsers)
End Sub

Private Sub grdUsers_Click()
    grdUsers.ToolTipText = "Double-click to view payroll"
End Sub

Private Sub grdUsers_DblClick()
    frmUserPayroll.Show
End Sub
