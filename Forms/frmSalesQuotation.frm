VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSalesQuotation 
   BackColor       =   &H80000005&
   Caption         =   "Sales Quotation"
   ClientHeight    =   9375
   ClientLeft      =   4410
   ClientTop       =   465
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   10590
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
      Left            =   6840
      TabIndex        =   3
      Text            =   "Search "
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton btnSalesQuoteAdd 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdSalesQuote 
      Height          =   8175
      Left            =   240
      TabIndex        =   0
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
      Caption         =   "Sales Quotations"
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
      TabIndex        =   2
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "frmSalesQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IsPurchased As Boolean
Private ReferenceID As String

Private Sub btnSalesQuoteAdd_Click()
    frmSalesQuotationAddSales.Show
End Sub

Private Sub Form_Load()
    IsPurchased = False
    
    Call Quotation.DisplayGrid(grdSalesQuote)
End Sub

Private Sub grdSalesQuote_DblClick()
    With grdSalesQuote
        ReferenceID = .TextMatrix(.Row, 1)
        IsPurchased = True
    End With
    
    query = "INSERT INTO tbl_tp_quotation_meta (meta_id, quotation_number, is_purchased) VALUES ('','" & ReferenceID & "','" & IsPurchased & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmSalesQuotation.Caption
End Sub
