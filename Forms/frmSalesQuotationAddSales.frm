VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmSalesQuotationAddSales 
   Caption         =   "Add New Sales Quotation"
   ClientHeight    =   9375
   ClientLeft      =   1905
   ClientTop       =   390
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   10605
   Begin VB.ComboBox cboSalesType 
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
      ItemData        =   "frmSalesQuotationAddSales.frx":0000
      Left            =   7560
      List            =   "frmSalesQuotationAddSales.frx":000D
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton btnSalesQuoteSubTotal 
      Caption         =   "Total Price"
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
      Left            =   2400
      TabIndex        =   32
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton btnSalesQuotationAdd 
      Caption         =   "Add Quote"
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
      Left            =   8760
      TabIndex        =   15
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtSalesFreight 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtSalesPrice 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtSalesTax 
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3960
      Width           =   2775
   End
   Begin VB.TextBox txtSalesProfit 
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
      Left            =   2520
      TabIndex        =   7
      Top             =   4440
      Width           =   2775
   End
   Begin VB.TextBox txtSalesOpEx 
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
      Left            =   2520
      TabIndex        =   8
      Top             =   4920
      Width           =   2775
   End
   Begin VB.TextBox txtSalesRep 
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
      Left            =   2520
      TabIndex        =   9
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txtSalesTraining 
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
      Left            =   2520
      TabIndex        =   10
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox txtSalesMisc 
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
      Left            =   2520
      TabIndex        =   11
      Top             =   6360
      Width           =   2775
   End
   Begin VB.TextBox txtSalesDelivery 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   6840
      Width           =   2775
   End
   Begin VB.TextBox txtSalesOther 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   7320
      Width           =   2775
   End
   Begin VB.ComboBox cboSalesEquipment 
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
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cboSalesClient 
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reference Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   16
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtSalesNumber 
         Height          =   405
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComCtl2.DTPicker dtpSalesDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Format          =   16580609
      CurrentDate     =   41675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Type"
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
      Left            =   5640
      TabIndex        =   33
      Top             =   1560
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Taxes and Duties"
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
      TabIndex        =   31
      Top             =   3960
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add New Quotation"
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
      TabIndex        =   30
      Top             =   240
      Width           =   2565
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Net Price"
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
      TabIndex        =   29
      Top             =   3000
      Width           =   795
   End
   Begin VB.Label Label6 
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
      TabIndex        =   28
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Client"
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
      TabIndex        =   27
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Equipment"
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
      TabIndex        =   26
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date Created"
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
      TabIndex        =   25
      Top             =   1560
      Width           =   1155
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Gross Profit"
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
      TabIndex        =   24
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Operational Expense"
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
      TabIndex        =   23
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Representation"
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
      TabIndex        =   22
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Training"
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
      TabIndex        =   21
      Top             =   5880
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "AVR and UPS"
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
      TabIndex        =   20
      Top             =   6360
      Width           =   1155
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Delivery Charge"
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
      TabIndex        =   19
      Top             =   6840
      Width           =   1380
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Others"
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
      TabIndex        =   18
      Top             =   7320
      Width           =   585
   End
   Begin VB.Label lblSalesAmount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
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
      Left            =   3720
      TabIndex        =   17
      Top             =   7920
      Width           =   1530
   End
End
Attribute VB_Name = "frmSalesQuotationAddSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSalesQuotationAdd_Click()
    query = "INSERT INTO tbl_tp_quotations(ID, q_ref_num, q_date, q_type, q_equipment, q_client, q_price, q_freight, q_tax, q_profit, q_opex, q_rep, q_training, q_misc, q_delivery, q_other, total_amount) " & _
            "VALUES ('','" & Me.txtSalesNumber.Text & "','" & Me.dtpSalesDate.value & "','" & LCase(Me.cboSalesType.Text) & "','" & Me.cboSalesEquipment.Text & "','" & Me.cboSalesClient.Text & "','" & Me.txtSalesPrice.Text & "','" & Me.txtSalesFreight.Text & "','" & Me.txtSalesTax.Text & "','" & Me.txtSalesProfit.Text & "','" & Me.txtSalesOpEx.Text & "','" & Me.txtSalesRep.Text & "','" & Me.txtSalesTraining.Text & "','" & Me.txtSalesMisc.Text & "','" & Me.txtSalesDelivery.Text & "','" & Me.txtSalesOther & "','" & Me.lblSalesAmount & "')"
            
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmSalesQuotationAddSales.Caption
    
    Unload Me
End Sub

Private Sub btnSalesQuoteSubTotal_Click()
    Me.lblSalesAmount.Caption = (CDbl(Me.txtSalesPrice) + CDbl(Me.txtSalesFreight) + CDbl(Me.txtSalesTax) + CDbl(Me.txtSalesProfit) + CDbl(Me.txtSalesOpEx) + CDbl(Me.txtSalesRep) + CDbl(Me.txtSalesTraining) + CDbl(Me.txtSalesMisc) + CDbl(Me.txtSalesDelivery) + CDbl(Me.txtSalesOther))
End Sub

Private Sub cboSalesEquipment_LostFocus()
    Set Records = Connect.Execute("SELECT freight_cost FROM tbl_tp_products WHERE equipment_name = '" & cboSalesEquipment.Text & "'")
    Me.txtSalesFreight.Text = Records!freight_cost
End Sub

Private Sub Form_Load()
    Call Client.DisplayDropdown(cboSalesClient)
    Call Product.DisplayEquipment(cboSalesEquipment)
    
    Call Quotation.ToDouble(Me.txtSalesPrice)
    Call Quotation.ToDouble(Me.txtSalesFreight)
    Call Quotation.ToDouble(Me.txtSalesTax)
    Call Quotation.ToDouble(Me.txtSalesProfit)
    Call Quotation.ToDouble(Me.txtSalesOpEx)
    Call Quotation.ToDouble(Me.txtSalesRep)
    Call Quotation.ToDouble(Me.txtSalesTraining)
    Call Quotation.ToDouble(Me.txtSalesMisc)
    Call Quotation.ToDouble(Me.txtSalesDelivery)
    Call Quotation.ToDouble(Me.txtSalesOther)
End Sub

Private Sub txtSalesDelivery_Change()
    Call Quotation.ToDouble(Me.txtSalesDelivery)
End Sub

Private Sub txtSalesFreight_Change()
    Call Quotation.ToDouble(Me.txtSalesFreight)
End Sub

Private Sub txtSalesMisc_Change()
    Call Quotation.ToDouble(Me.txtSalesMisc)
End Sub

Private Sub txtSalesOpEx_Change()
    Call Quotation.ToDouble(Me.txtSalesOpEx)
End Sub

Private Sub txtSalesOther_Change()
    Call Quotation.ToDouble(Me.txtSalesOther)
End Sub

Private Sub txtSalesPrice_Change()
    Call Quotation.ToDouble(Me.txtSalesPrice)
    
    Me.txtSalesTax.Text = Tax * (CDbl(Me.txtSalesPrice))
    Me.txtSalesProfit.Text = Profit * (CDbl(Me.txtSalesPrice))
    Me.txtSalesOpEx.Text = OpEx * (CDbl(Me.txtSalesPrice))
    Me.txtSalesRep.Text = Rep * (CDbl(Me.txtSalesPrice))
End Sub

Private Sub txtSalesProfit_Change()
    Call Quotation.ToDouble(Me.txtSalesProfit)
End Sub

Private Sub txtSalesRep_Change()
    Call Quotation.ToDouble(Me.txtSalesRep)
End Sub

Private Sub txtSalesTax_Change()
    Call Quotation.ToDouble(Me.txtSalesTax)
End Sub

Private Sub txtSalesTraining_Change()
    Call Quotation.ToDouble(Me.txtSalesTraining)
End Sub
