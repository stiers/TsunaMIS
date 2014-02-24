VERSION 5.00
Begin VB.Form frmUserGroups 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "User Groups"
   ClientHeight    =   8085
   ClientLeft      =   5625
   ClientTop       =   1530
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   Begin VB.ListBox lstUserGroup 
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
      TabIndex        =   2
      Top             =   1440
      Width           =   6015
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
      ItemData        =   "frmUserGroups.frx":0000
      Left            =   240
      List            =   "frmUserGroups.frx":000D
      TabIndex        =   1
      Text            =   "Bulk Action"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton btnGroupBulkAction 
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
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblClients 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "User Groups"
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
      Width           =   1560
   End
End
Attribute VB_Name = "frmUserGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
