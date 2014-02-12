VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmAccExpType 
   Caption         =   "Miscellaneous"
   ClientHeight    =   5655
   ClientLeft      =   4305
   ClientTop       =   1890
   ClientWidth     =   7230
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmAccExpType.frx":0000
   ScaleHeight     =   0
   ScaleWidth      =   0
   Begin VB.ComboBox Combo1 
      Height          =   435
      ItemData        =   "frmAccExpType.frx":0342
      Left            =   2160
      List            =   "frmAccExpType.frx":034C
      TabIndex        =   10
      Top             =   1920
      Width           =   4815
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   4815
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   4
      Top             =   2520
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   41682
   End
   Begin VB.CommandButton btnExpenseType 
      Caption         =   "Save"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      Height          =   315
      Left            =   480
      TabIndex        =   9
      Top             =   1920
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   315
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice #:"
      Height          =   315
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      Height          =   435
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Miscellaneous:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmAccExpType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


