VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEngineerVEH 
   BackColor       =   &H80000005&
   Caption         =   "Equipment History"
   ClientHeight    =   9375
   ClientLeft      =   360
   ClientTop       =   1260
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   15120
   Begin VB.CommandButton btnServiceQuotationAdd 
      Caption         =   "Back to Reports"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdSalesQuote 
      Height          =   8175
      Left            =   240
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
      Caption         =   "Equipment History"
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2400
   End
End
Attribute VB_Name = "frmEngineerVEH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
