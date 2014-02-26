VERSION 5.00
Begin VB.Form Clients 
   BackColor       =   &H80000005&
   Caption         =   "Clients"
   ClientHeight    =   8580
   ClientLeft      =   4035
   ClientTop       =   2235
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8580
   ScaleWidth      =   10110
   Begin VB.ListBox ListBox_Clients 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1680
      Width           =   4575
   End
   Begin VB.ComboBox ComboBox_ClientBulkAction 
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
      ItemData        =   "frmClients.frx":0000
      Left            =   240
      List            =   "frmClients.frx":000D
      TabIndex        =   2
      Text            =   "Bulk Action"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton CommandButton_ClientApplyAction 
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
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CommandButton_ClientAdd 
      BackColor       =   &H8000000D&
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
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label_ClientMobile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientMobile"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   7800
      Width           =   1635
   End
   Begin VB.Label Label_ClientFax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientFax"
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
      Left            =   5280
      TabIndex        =   17
      Top             =   6840
      Width           =   1305
   End
   Begin VB.Label Label_ClientTelephone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientTelephone"
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
      Left            =   5280
      TabIndex        =   16
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label_ClientMail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientMail"
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
      Left            =   5280
      TabIndex        =   15
      Top             =   4920
      Width           =   1395
   End
   Begin VB.Label Label_ClientAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientAddress"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label Label_ClientCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientCompany"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   3000
      Width           =   1845
   End
   Begin VB.Label Label_ClientPosition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label_ClientPosition"
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
      Left            =   5280
      TabIndex        =   12
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   6480
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
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
      Left            =   5280
      TabIndex        =   9
      Top             =   5520
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail"
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
      Left            =   5280
      TabIndex        =   8
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Address"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   3600
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label1 
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
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label_Clients 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Clients"
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
      TabIndex        =   4
      Top             =   480
      Width           =   870
   End
End
Attribute VB_Name = "Clients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_ClientAdd_Click()
    ClientAdd.Show
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    'from TsunaClients class
    Call Client.DisplayClientNameAsList(ListBox_Clients)
End Sub

Private Sub ListBox_Clients_Click()
    id = Me.ListBox_Clients.ListIndex + 1
    
    If Record.State = 1 Then Record.Close
    
    query = "SELECT client_position, client_company, client_address, client_email, client_telephone, client_fax, client_mobile FROM tbl_tp_clients WHERE ID = '" & id & "'"
    
    Record.Open query, Connect
    
    Me.Label_ClientPosition.Caption = Record!client_position
    Me.Label_ClientCompany.Caption = Record!client_company
    Me.Label_ClientAddress.Caption = Record!client_address
    Me.Label_ClientMail.Caption = Record!client_email
    Me.Label_ClientTelephone.Caption = Record!client_telephone
    Me.Label_ClientFax.Caption = Record!client_fax
    Me.Label_ClientMobile.Caption = Record!client_mobile
End Sub
