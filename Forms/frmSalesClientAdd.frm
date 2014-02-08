VERSION 5.00
Begin VB.Form frmSalesClientAdd 
   Caption         =   "Add New Client"
   ClientHeight    =   8085
   ClientLeft      =   29025
   ClientTop       =   3585
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   Begin VB.CommandButton btnSalesClientAdd 
      Caption         =   "Add Client"
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
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtClientName 
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
      Top             =   2760
      Width           =   2775
   End
   Begin VB.TextBox txtClientAddress 
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
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Client Address"
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
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Client Name"
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
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add New Client"
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
      Width           =   2010
   End
End
Attribute VB_Name = "frmSalesClientAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalesClientAdd_Click()
    query = "INSERT INTO tbl_tp_clients(client_id, client_name, client_address) " & _
            "VALUES ('', '" & Me.txtClientName & "', '" & Me.txtClientAddress & "')"
            
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmSalesClientAdd.Caption
    
    Unload Me
End Sub

