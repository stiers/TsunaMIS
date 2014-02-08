VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLogisticsPL 
   BackColor       =   &H80000005&
   Caption         =   "Product Line"
   ClientHeight    =   8070
   ClientLeft      =   2790
   ClientTop       =   1485
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid grdProducts 
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   12303
      _Version        =   393216
      AllowUserResizing=   2
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
   Begin VB.CommandButton btnLogisticsPLAdd 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Product Line"
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
      Width           =   1605
   End
End
Attribute VB_Name = "frmLogisticsPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogisticsPLAdd_Click()
    frmLogisticsPLAdd.Show
End Sub

Private Sub Form_Load()
    Call Product.DisplayGrid(grdProducts)
End Sub
