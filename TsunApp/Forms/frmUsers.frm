VERSION 5.00
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   8085
   ClientLeft      =   5175
   ClientTop       =   1875
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
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
