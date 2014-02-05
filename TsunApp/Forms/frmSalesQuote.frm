VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesQuote 
   Caption         =   "Sales Quotation"
   ClientHeight    =   8535
   ClientLeft      =   5400
   ClientTop       =   1560
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   6420
   Begin VB.Frame fraSalesNum 
      Caption         =   "Reference Number"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   30
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtSalesNum 
         Height          =   405
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComCtl2.DTPicker dteSalesDate 
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   1440
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
      Format          =   41943041
      CurrentDate     =   41675
   End
   Begin VB.ComboBox cboSalesSupplier 
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
      Left            =   3360
      TabIndex        =   28
      Top             =   2400
      Width           =   2775
   End
   Begin VB.ComboBox cboSalesPersonnel 
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
      Left            =   3360
      TabIndex        =   27
      Top             =   1920
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
      Left            =   3360
      TabIndex        =   24
      Top             =   7200
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
      Left            =   3360
      TabIndex        =   23
      Top             =   6720
      Width           =   2775
   End
   Begin VB.TextBox txtSalesAddOn 
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
      Left            =   3360
      TabIndex        =   22
      Top             =   6240
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
      Left            =   3360
      TabIndex        =   21
      Top             =   5760
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
      Left            =   3360
      TabIndex        =   20
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox txtSalesExpense 
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
      Left            =   3360
      TabIndex        =   19
      Top             =   4800
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
      Left            =   3360
      TabIndex        =   12
      Top             =   4320
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
      Left            =   3360
      TabIndex        =   10
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtSalesNetPrice 
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
      Left            =   3360
      TabIndex        =   2
      Top             =   2880
      Width           =   2775
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
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.CommandButton btnAddQuote 
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
      Left            =   240
      TabIndex        =   0
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label lblSalesTotalPrice 
      AutoSize        =   -1  'True
      Caption         =   "Total Price"
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
      Left            =   3360
      TabIndex        =   26
      Top             =   7800
      Width           =   915
   End
   Begin VB.Label lblSalesTotalAmount 
      AutoSize        =   -1  'True
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
      Left            =   5520
      TabIndex        =   25
      Top             =   7800
      Width           =   570
   End
   Begin VB.Label lblSalesOther 
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
      Top             =   7200
      Width           =   585
   End
   Begin VB.Label lblSalesDelivery 
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
      TabIndex        =   17
      Top             =   6720
      Width           =   1380
   End
   Begin VB.Label lblSalesAddOn 
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
      TabIndex        =   16
      Top             =   6240
      Width           =   1155
   End
   Begin VB.Label lblSalesTraining 
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
      TabIndex        =   15
      Top             =   5760
      Width           =   705
   End
   Begin VB.Label lblSalesRep 
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
      TabIndex        =   14
      Top             =   5280
      Width           =   1320
   End
   Begin VB.Label lblSalesExpense 
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
      TabIndex        =   13
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblSalesProfit 
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
      TabIndex        =   11
      Top             =   4320
      Width           =   1035
   End
   Begin VB.Label lblSalesDate 
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
      TabIndex        =   9
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label lblSalesPersonnel 
      AutoSize        =   -1  'True
      Caption         =   "Personnel"
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
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblSalesSupplier 
      AutoSize        =   -1  'True
      Caption         =   "Supplier"
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
      TabIndex        =   7
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label lblSalesFreight 
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
      TabIndex        =   6
      Top             =   3360
      Width           =   1050
   End
   Begin VB.Label lblSalesNetPrice 
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
      TabIndex        =   5
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label lblUserAddNew 
      AutoSize        =   -1  'True
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
      TabIndex        =   4
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label lblSalesTax 
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
      TabIndex        =   3
      Top             =   3840
      Width           =   1485
   End
End
Attribute VB_Name = "frmSalesQuote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddQuote_Click()
    Dim ExcelApp As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet
    Dim ExcelName As String
    
    Set ExcelApp = New Excel.Application
    Set ExcelWB = ExcelApp.Workbooks.Add
    Set ExcelWS = ExcelWB.Worksheets.Add

    ExcelName = Me.txtSalesNum.Text & ".xlsx"
    
    query = "INSERT INTO tbl_quotations(q_id, q_num, q_type) " & _
            "VALUES ('', '" & Me.txtSalesNum.Text & "', 'sales')"
    
    ExcelWS.Cells(3, 2) = Me.dteSalesDate.value
    ExcelWS.Cells(4, 2) = Me.cboSalesPersonnel.Text
    ExcelWS.Cells(5, 2) = Me.cboSalesSupplier.Text
    ExcelWS.Cells(6, 2) = Me.txtSalesNetPrice.Text
    ExcelWS.Cells(7, 2) = Me.txtSalesFreight.Text
    ExcelWS.Cells(8, 2) = Me.txtSalesTax.Text
    ExcelWS.Cells(9, 2) = Me.txtSalesProfit.Text
    ExcelWS.Cells(10, 2) = Me.txtSalesExpense.Text
    ExcelWS.Cells(11, 2) = Me.txtSalesRep.Text
    ExcelWS.Cells(12, 2) = Me.txtSalesTraining.Text
    ExcelWS.Cells(13, 2) = Me.txtSalesAddOn.Text
    ExcelWS.Cells(14, 2) = Me.txtSalesDelivery.Text
    ExcelWS.Cells(15, 2) = Me.txtSalesOther.Text
    ExcelWS.Cells(17, 2) = Me.lblSalesTotalAmount
    
    Connect.Execute query
    
    ExcelWB.SaveAs (App.Path & "\Files\Workbooks\" & ExcelName & "")
    ExcelWB.Close (SaveChanges = True)
    
    Set ExcelApp = Nothing
End Sub

Private Sub Form_Load()
    Call Users.Dropdown(cboSalesPersonnel)
    Call Suppliers.Dropdown(cboSalesSupplier)
End Sub
