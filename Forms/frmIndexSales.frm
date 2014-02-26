VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form IndexSales 
   BackColor       =   &H80000005&
   Caption         =   "Sales"
   ClientHeight    =   9345
   ClientLeft      =   585
   ClientTop       =   1530
   ClientWidth     =   16500
   Icon            =   "frmIndexSales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   16500
   Begin MSFlexGridLib.MSFlexGrid grdSalesQuote 
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   5318
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
   Begin MSFlexGridLib.MSFlexGrid grdServiceQuote 
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   5953
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   8880
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Service Quotations"
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
      TabIndex        =   4
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome user! You have no pending quotations at this moment"
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
      Left            =   10560
      TabIndex        =   2
      Top             =   240
      Width           =   5580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   1
      Top             =   240
      Width           =   2145
   End
   Begin VB.Menu mnuQuote 
      Caption         =   "Quotations"
      Begin VB.Menu mnuQuoteRequest 
         Caption         =   "Request for Quotation"
      End
      Begin VB.Menu mnuQuoteAddSales 
         Caption         =   "Add Sales Quotation"
      End
      Begin VB.Menu mnuQuoteAddSvc 
         Caption         =   "Add Service Quotation"
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "Project Bank"
   End
End
Attribute VB_Name = "IndexSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'from Grid Header module
    Call SalesQuotationHead
    Call Sales.DisplaySalesQuotesToGrid(grdSalesQuote)
    
    'from Grid Header module
    Call ServiceQuotationHead
    Call Sales.DisplayServiceQuotesToGrid(grdServiceQuote)
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.Label_SystemDate.Caption = Format$(Now, "m/d/yy")
    Me.Label_SystemTime.Caption = Format$(Now, "hh:mm AM/PM")
End Sub

Private Sub grdSalesQuote_DblClick()
    SalesQuotation.Show
    
    With SalesQuotation
        If Record.State = 1 Then Record.Close
    
        query = "SELECT * FROM tbl_tp_quotations WHERE q_ref_num = '" & grdSalesQuote.TextMatrix(grdSalesQuote.Row, 1) & "'"
        
        Record.Open query, Connect
        
        .btnSalesQuotationAdd.Visible = False
        .btnConvertToSalesInvoice.Visible = True
        .CommandButton_ClientUpdate.Visible = True
        .CommandButton_ClientDelete.Visible = True
        .CommandButton_ClientPrint.Visible = True
    
        CurrentQuotationID = Record!ID  'check the Core Module for usage
        .txtSalesNumber.Text = Record!q_ref_num
        .cboSalesType.Text = Record!q_type
        .dtpSalesDate.value = Record!q_date
        .cboSalesEquipment.Text = Record!q_equipment
        .cboSalesClient.Text = Record!q_client
        .txtSalesPrice.Text = Format(Record!q_price, "0.00")
        .txtSalesFreight.Text = Format(Record!q_freight, "0.00")
        .txtSalesTax.Text = Format(Record!q_tax, "0.00")
        .txtSalesProfit.Text = Format(Record!q_profit, "0.00")
        .txtSalesOpEx.Text = Format(Record!q_opex, "0.00")
        .txtSalesRep.Text = Format(Record!q_rep, "0.00")
        .txtSalesTraining.Text = Format(Record!q_training, "0.00")
        .txtSalesMisc.Text = Format(Record!q_misc, "0.00")
        .txtSalesDelivery.Text = Format(Record!q_delivery, "0.00")
        .txtSalesOther.Text = Format(Record!q_other, "0.00")
        .lblSalesAmount.Caption = Format(Record!total_amount, "0.00")
    End With
End Sub

Private Sub grdServiceQuote_DblClick()
    SalesQuotation.Show
    
    With SalesQuotation
        If Record.State = 1 Then Record.Close
    
        query = "SELECT * FROM tbl_tp_quotations WHERE q_ref_num = '" & grdServiceQuote.TextMatrix(grdServiceQuote.Row, 1) & "'"
    
        Record.Open query, Connect
        
        .btnSalesQuotationAdd.Visible = False
        .btnConvertToSalesInvoice.Visible = True
        .CommandButton_ClientUpdate.Visible = True
        .CommandButton_ClientDelete.Visible = True
        .CommandButton_ClientPrint.Visible = True
        
        CurrentQuotationID = Record!ID  'check the Core Module for usage
        .txtSalesNumber.Text = Record!q_ref_num
        .cboSalesType.Text = Record!q_type
        .dtpSalesDate.value = Record!q_date
        .cboSalesEquipment.Text = Record!q_equipment
        .cboSalesClient.Text = Record!q_client
        .txtSalesPrice.Text = Record!q_price
        .txtSalesFreight.Text = Record!q_freight
        .txtSalesTax.Text = Record!q_tax
        .txtSalesProfit.Text = Record!q_profit
        .txtSalesOpEx.Text = Record!q_opex
        .txtSalesRep.Text = Record!q_rep
        .txtSalesTraining.Text = Record!q_training
        .txtSalesMisc.Text = Record!q_misc
        .txtSalesDelivery.Text = Record!q_delivery
        .txtSalesOther.Text = Record!q_other
        .lblSalesAmount.Caption = Record!total_amount
    End With
End Sub

Private Sub mnuQuoteAddSales_Click()
    SalesQuotation.Show
    SalesQuotation.cboSalesType.Text = "Sales"
    
    SalesQuotation.CommandButton_ClientUpdate.Visible = False
    SalesQuotation.CommandButton_ClientDelete.Visible = False
End Sub

Private Sub mnuQuoteAddSvc_Click()
    SalesQuotation.Show
    SalesQuotation.cboSalesType.Text = "Maintenance"
    
    SalesQuotation.CommandButton_ClientUpdate.Visible = False
    SalesQuotation.CommandButton_ClientDelete.Visible = False
End Sub
