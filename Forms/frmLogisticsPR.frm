VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLogisticsPR 
   BackColor       =   &H80000005&
   Caption         =   "Purchase Request"
   ClientHeight    =   9375
   ClientLeft      =   3840
   ClientTop       =   1155
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   10605
   Begin MSFlexGridLib.MSFlexGrid grdPurchase 
      Height          =   8175
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   10095
      _ExtentX        =   17806
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
      TabIndex        =   0
      Top             =   240
      Width           =   2265
   End
End
Attribute VB_Name = "frmLogisticsPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Quotation.DisplayRequest(grdPurchase)
End Sub

Private Sub grdPurchase_DblClick()
    With grdPurchase
        frmLogisticsPRAdd.txtQuotationNumber.Text = .TextMatrix(.Row, 1)
        frmLogisticsPRAdd.dtpPurchaseDate.value = .TextMatrix(.Row, 2)
        frmLogisticsPRAdd.txtPurchaseEq.Text = .TextMatrix(.Row, 3)
        frmLogisticsPRAdd.txtPurchaseClient.Text = .TextMatrix(.Row, 4)
        frmLogisticsPRAdd.txtPurchaseAdd.Text = .TextMatrix(.Row, 5)
        frmLogisticsPRAdd.txtPurchasePrice.Text = .TextMatrix(.Row, 6)
    End With
    
    frmLogisticsPRAdd.Show
End Sub
