VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEngineerDAR 
   BackColor       =   &H80000005&
   Caption         =   "Daily Activity Record"
   ClientHeight    =   9405
   ClientLeft      =   210
   ClientTop       =   420
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   15075
   Begin VB.TextBox txtSearch 
      Alignment       =   1  'Right Justify
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
      Left            =   11280
      TabIndex        =   3
      Text            =   "Search "
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton btnServiceQuotationAdd 
      Caption         =   "Add Report"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdSalesQuote 
      Height          =   8175
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   14420
      _Version        =   393216
      AllowUserResizing=   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Daily Activity Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2700
   End
End
Attribute VB_Name = "frmEngineerDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnServiceQuotationAdd_Click()
    frmEngineerDARAdd.Show vbModal, Me
End Sub
