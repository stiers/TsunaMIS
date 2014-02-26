VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SalesQuotation 
   BackColor       =   &H80000005&
   Caption         =   "Sales Quotation"
   ClientHeight    =   8550
   ClientLeft      =   3975
   ClientTop       =   960
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10080
   Begin VB.CommandButton btnConvertToSalesInvoice 
      Caption         =   "Convert to Sales Invoice"
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
      Left            =   2520
      TabIndex        =   37
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton CommandButton_ClientPrint 
      BackColor       =   &H80000005&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CommandButton_ClientDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton CommandButton_ClientUpdate 
      BackColor       =   &H0000FF00&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton btnSalesQuotationAdd 
      BackColor       =   &H8000000D&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reference Number"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   6840
      TabIndex        =   19
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox cboSalesEquipment 
      Appearance      =   0  'Flat
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
      Left            =   2280
      TabIndex        =   2
      Top             =   2280
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   7080
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
      Left            =   2280
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   6600
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
      Left            =   2280
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   6120
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
      Left            =   2280
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   5640
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
      Left            =   2280
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   5160
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   4680
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   4200
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3720
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   2760
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   3240
      Width           =   2535
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
      Left            =   2280
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
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
      ItemData        =   "frmSalesQuotation.frx":0000
      Left            =   6840
      List            =   "frmSalesQuotation.frx":000D
      TabIndex        =   15
      Top             =   1320
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtpSalesDate 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
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
      Format          =   108134401
      CurrentDate     =   41675
   End
   Begin VB.Label lblSalesAmount 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   3480
      TabIndex        =   35
      Top             =   7680
      Width           =   1530
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   34
      Top             =   7080
      Width           =   585
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   33
      Top             =   6600
      Width           =   1380
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   32
      Top             =   6120
      Width           =   1155
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   31
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   30
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   29
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   28
      Top             =   4200
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   27
      Top             =   1320
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Top             =   2280
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   25
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   24
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   23
      Top             =   2760
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Quotation"
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
      TabIndex        =   22
      Top             =   480
      Width           =   2025
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   21
      Top             =   3720
      Width           =   1485
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   5400
      TabIndex        =   20
      Top             =   1320
      Width           =   420
   End
End
Attribute VB_Name = "SalesQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tax, Profit, OpEx, Rep As Double

Private Sub btnConvertToSalesInvoice_Click()
    SalesInvoice.Show
End Sub

Private Sub btnSalesQuotationAdd_Click()
    Call Sales.AddSales(Me.txtSalesNumber.Text, Me.dtpSalesDate.value, Me.cboSalesType.Text, Me.cboSalesEquipment.Text, Me.cboSalesClient.Text, Me.txtSalesPrice.Text, Me.txtSalesFreight.Text, Me.txtSalesTax.Text, Me.txtSalesProfit.Text, Me.txtSalesOpEx.Text, Me.txtSalesRep.Text, Me.txtSalesTraining.Text, Me.txtSalesMisc.Text, Me.txtSalesDelivery.Text, Me.txtSalesOther.Text, Me.lblSalesAmount.Caption)
    Unload Me
End Sub

Private Sub btnSalesQuoteSubTotal_Click()
    Me.lblSalesAmount.Caption = (CDbl(Me.txtSalesPrice) + CDbl(Me.txtSalesFreight) + CDbl(Me.txtSalesTax) + CDbl(Me.txtSalesProfit) + CDbl(Me.txtSalesOpEx) + CDbl(Me.txtSalesRep) + CDbl(Me.txtSalesTraining) + CDbl(Me.txtSalesMisc) + CDbl(Me.txtSalesDelivery) + CDbl(Me.txtSalesOther))
End Sub

Private Sub CommandButton_ClientDelete_Click()
    Call Sales.DeleteSales
    Unload Me
End Sub

Private Sub CommandButton_ClientUpdate_Click()
    Call Sales.UpdateSales(Me.txtSalesNumber.Text, Me.dtpSalesDate.value, Me.cboSalesType.Text, Me.cboSalesEquipment.Text, Me.cboSalesClient.Text, Me.txtSalesPrice.Text, Me.txtSalesFreight.Text, Me.txtSalesTax.Text, Me.txtSalesProfit.Text, Me.txtSalesOpEx.Text, Me.txtSalesRep.Text, Me.txtSalesTraining.Text, Me.txtSalesMisc.Text, Me.txtSalesDelivery.Text, Me.txtSalesOther.Text, Me.lblSalesAmount.Caption)
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    Me.CommandButton_ClientUpdate.Visible = False
    Me.CommandButton_ClientDelete.Visible = False
    Me.CommandButton_ClientPrint.Visible = False
    Me.btnConvertToSalesInvoice.Visible = False
    
    Me.dtpSalesDate.value = Format(Now, "yyyy" & "-" & "MM" & "-" & "dd")
    Me.lblSalesAmount.Caption = Format(dblTotal, "0.00")
    
     'from TsunaClients class
    Call Client.DisplayCompanyNameAsCombo(cboSalesClient)
    
    'from TsunaLogistics class
    Call Logistics.DisplayProductToCombo(cboSalesEquipment)
    
    'Global Variables
    Tax = 0.15
    Profit = 0.31
    OpEx = 0.14
    Rep = 0.02
End Sub

Private Sub txtSalesDelivery_Change()
    Call Sales.ToDouble(Me.txtSalesDelivery)
End Sub

Private Sub txtSalesFreight_Change()
    Call Sales.ToDouble(Me.txtSalesFreight)
End Sub

Private Sub txtSalesMisc_Change()
    Call Sales.ToDouble(Me.txtSalesMisc)
End Sub

Private Sub txtSalesOpEx_Change()
    Call Sales.ToDouble(Me.txtSalesOpEx)
End Sub

Private Sub txtSalesOther_Change()
    Call Sales.ToDouble(Me.txtSalesOther)
End Sub

Private Sub txtSalesPrice_Change()
    Call Sales.ToDouble(Me.txtSalesPrice)
    
    Me.txtSalesTax.Text = Tax * (CDbl(Me.txtSalesPrice))
    Me.txtSalesProfit.Text = Profit * (CDbl(Me.txtSalesPrice))
    Me.txtSalesOpEx.Text = OpEx * (CDbl(Me.txtSalesPrice))
    Me.txtSalesRep.Text = Rep * (CDbl(Me.txtSalesPrice))
End Sub

Private Sub txtSalesProfit_Change()
    Call Sales.ToDouble(Me.txtSalesProfit)
End Sub

Private Sub txtSalesRep_Change()
    Call Sales.ToDouble(Me.txtSalesRep)
End Sub

Private Sub txtSalesTax_Change()
    Call Sales.ToDouble(Me.txtSalesTax)
End Sub

Private Sub txtSalesTraining_Change()
    Call Sales.ToDouble(Me.txtSalesTraining)
End Sub
