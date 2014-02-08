VERSION 5.00
Begin VB.Form frmLogisticsPLAdd 
   Caption         =   "Add New Product"
   ClientHeight    =   8085
   ClientLeft      =   5865
   ClientTop       =   1890
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFreightCost 
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
      TabIndex        =   7
      Top             =   3240
      Width           =   2775
   End
   Begin VB.ComboBox cboProductName 
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
      Text            =   "Choose nothing"
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton btnProductClose 
      Caption         =   "&Close"
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
      Left            =   4800
      TabIndex        =   3
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtEquipmentName 
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
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton btnProductAdd 
      Caption         =   "&Add"
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
      TabIndex        =   2
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Freight Cost"
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
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add New Product"
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
      TabIndex        =   6
      Top             =   240
      Width           =   2265
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Product or Equipment"
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
      Top             =   2640
      Width           =   1920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Product Line"
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
      Top             =   3840
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogisticsPLAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnProductAdd_Click()
    Call Product.AddItem(Trim(Me.txtEquipmentName.Text), Me.txtFreightCost.Text, Me.cboProductName.Text)
End Sub

Private Sub btnProductClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Product.DisplayDropdown(cboProductName)
End Sub
