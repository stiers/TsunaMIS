VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQuotations 
   ClientHeight    =   8325
   ClientLeft      =   4230
   ClientTop       =   1665
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   9600
   Begin MSFlexGridLib.MSFlexGrid grdQuotation 
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
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
   Begin VB.Label lblUserAddNew 
      AutoSize        =   -1  'True
      Caption         =   "Quotations"
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
      Top             =   360
      Width           =   1425
   End
End
Attribute VB_Name = "frmQuotations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Quote.Display(grdQuotation)
End Sub
