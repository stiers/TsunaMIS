VERSION 5.00
Begin VB.Form frmSalesClient 
   BackColor       =   &H80000005&
   Caption         =   "Clients"
   ClientHeight    =   8085
   ClientLeft      =   29385
   ClientTop       =   2640
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   Begin VB.CommandButton btnClientAdd 
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
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton btnClientBulkAction 
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
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboClientBulkAction 
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
      ItemData        =   "frmSalesClient.frx":0000
      Left            =   240
      List            =   "frmSalesClient.frx":000D
      TabIndex        =   1
      Text            =   "Bulk Action"
      Top             =   840
      Width           =   1815
   End
   Begin VB.ListBox lstClients 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6330
      Left            =   240
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Label Label1 
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
      Top             =   240
      Width           =   870
   End
End
Attribute VB_Name = "frmSalesClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClientAdd_Click()
    frmSalesClientAdd.Show
End Sub

Private Sub Form_Load()
    Call Client.DisplayList(lstClients)
End Sub
