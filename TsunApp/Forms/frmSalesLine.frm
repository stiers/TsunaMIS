VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSalesLine 
   Caption         =   "Sales Line"
   ClientHeight    =   8085
   ClientLeft      =   4530
   ClientTop       =   1935
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9510
   Begin VB.CommandButton btnAddNewSalesLine 
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
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSalesLine 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9255
      _ExtentX        =   16325
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
   Begin VB.Label lblSalesLine 
      AutoSize        =   -1  'True
      Caption         =   "Sales Line"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "frmSalesLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SalesLine.Display(grdSalesLine)
End Sub
