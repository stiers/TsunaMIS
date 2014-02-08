VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H80000005&
   Caption         =   "Welcome"
   ClientHeight    =   7350
   ClientLeft      =   2715
   ClientTop       =   2730
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   2520
      Picture         =   "frmSecurity.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   2175
      TabIndex        =   4
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   4800
      Width           =   1035
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "TsunaMIS"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   750
      Left            =   4920
      TabIndex        =   6
      Top             =   2040
      Width           =   2220
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Tsuna Management Information System"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   5
      Top             =   2880
      Width           =   4365
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mUsername As String
Private mPassword As String
Private mCancel As Boolean

Private Sub cmdCancel_Click()
    mCancel = True
    
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mCancel = False
    mUsername = Me.txtUsername.Text
    mPassword = Me.txtPassword.Text
    
    Unload Me
End Sub

Public Function GetLogIn(ByRef Username As String, ByRef Password As String, Owner As Object) As Boolean
    Me.txtUsername.Text = Username
    
    Me.Show vbModal, Owner
    
    Username = mUsername
    Password = mPassword
    
    GetLogIn = Not mCancel
End Function

Private Sub Form_Activate()
    If Len(Me.txtUsername.Text) > 0 Then Me.txtPassword.SetFocus
End Sub

Private Sub txtUsername_GotFocus()
    txtUsername.SelStart = 0
    txtUsername.SelLength = Len(txtUsername.Text)
End Sub
