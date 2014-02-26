VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseRequest 
   BackColor       =   &H80000005&
   Caption         =   "Purchase Request"
   ClientHeight    =   8565
   ClientLeft      =   4155
   ClientTop       =   1230
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8565
   ScaleWidth      =   10080
   Begin VB.CommandButton btnReqDelete 
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton btnReqUpdate 
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton btnReqAdd 
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton btnReqPrint 
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
      TabIndex        =   25
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox Text_ReqClientNumber 
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
      Left            =   7200
      TabIndex        =   23
      Top             =   5880
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqClientMobile 
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
      TabIndex        =   21
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqClientDeliveryAdd 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqEqProdNum 
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
      TabIndex        =   17
      Top             =   5880
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqEqModel 
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
      TabIndex        =   15
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqEqBrand 
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
      Top             =   4680
      Width           =   2415
   End
   Begin VB.ComboBox Combo_ClientName 
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
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqEqName 
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
      Top             =   4080
      Width           =   2415
   End
   Begin VB.TextBox Text_ReqNumber 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text_InvNumber 
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
      Left            =   7200
      TabIndex        =   0
      Top             =   2040
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker_DateIssued 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
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
   Begin MSComCtl2.DTPicker DTPicker_DateDelivery 
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number"
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
      Left            =   5160
      TabIndex        =   24
      Top             =   5880
      Width           =   1440
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
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
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Address"
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
      Left            =   5160
      TabIndex        =   20
      Top             =   4080
      Width           =   1470
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Number"
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
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Top             =   5280
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Brand"
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
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Issued"
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
      TabIndex        =   12
      Top             =   1320
      Width           =   1020
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
      TabIndex        =   11
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label5 
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
      TabIndex        =   10
      Top             =   4080
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      TabIndex        =   9
      Top             =   480
      Width           =   2265
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Number"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   2760
      Width           =   1650
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Number"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   2040
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Delivery"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   1320
      Width           =   1395
   End
End
Attribute VB_Name = "PurchaseRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnReqAdd_Click()
    Call Logistics.AddPurchase(Me.DTPicker_DateIssued.value, Me.DTPicker_DateDelivery.value, Me.Text_ReqNumber.Text, Me.Text_InvNumber.Text, Me.Combo_ClientName.Text, Me.Text_ReqClientDeliveryAdd.Text, Me.Text_ReqClientNumber.Text, Me.Text_ReqClientMobile.Text, Me.Text_ReqEqName.Text, Me.Text_ReqEqBrand.Text, Me.Text_ReqEqModel.Text, Me.Text_ReqEqProdNum.Text)
    Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    
    With SalesInvoice
        Me.DTPicker_DateIssued.value = .DTPicker_DueDate.value
        Me.Combo_ClientName.Text = .Combo_QuoteClient.Text
        Me.Text_InvNumber.Text = .Text_InvNum.Text
        Me.Text_ReqEqName.Text = .Combo_QuoteEq.Text
    End With
End Sub
