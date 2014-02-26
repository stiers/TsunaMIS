VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IndexLogistics 
   BackColor       =   &H80000005&
   Caption         =   "Logistics"
   ClientHeight    =   9330
   ClientLeft      =   630
   ClientTop       =   1335
   ClientWidth     =   16485
   LinkTopic       =   "Form1"
   ScaleHeight     =   9330
   ScaleWidth      =   16485
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
      Left            =   12720
      TabIndex        =   1
      Text            =   "Search "
      Top             =   480
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
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grdDailyActivityRecord 
      Height          =   6735
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   11880
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
   Begin VB.Label Label_SystemDate 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label_SystemDate"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   13440
      TabIndex        =   5
      Top             =   8880
      Width           =   1365
   End
   Begin VB.Label Label_SystemTime 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label_SystemTime"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   14880
      TabIndex        =   4
      Top             =   8880
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Purchase Request"
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
      Top             =   360
      Width           =   2265
   End
End
Attribute VB_Name = "IndexLogistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnServiceQuotationAdd_Click()
    PurchaseRequest.Show
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.Label_SystemDate.Caption = Format$(Now, "m/d/yy")
    Me.Label_SystemTime.Caption = Format$(Now, "hh:mm AM/PM")
    
    Call PurchaseRequestHead
    Call Logistics.DisplayPurchaseRequests(grdDailyActivityRecord)
End Sub

Private Sub mnuProductAdd_Click()
    Products.Show
End Sub
