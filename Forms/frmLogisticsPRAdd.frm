VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogisticsPRAdd 
   Caption         =   "Print Purchase Request"
   ClientHeight    =   9390
   ClientLeft      =   3075
   ClientTop       =   1080
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   10575
   Begin VB.Frame Frame2 
      Caption         =   "Quotation Number"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   14
      Top             =   1320
      Width           =   2775
      Begin VB.TextBox txtQuotationNumber 
         Height          =   405
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Request Number"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      Begin VB.TextBox txtPurchaseNumber 
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.TextBox txtPurchasePrice 
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
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtPurchaseAdd 
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
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtPurchaseEq 
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
   Begin VB.TextBox txtPurchaseClient 
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
      TabIndex        =   1
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton btnPurchaseAdd 
      Caption         =   "Print Request"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtpPurchaseDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
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
      Format          =   113901569
      CurrentDate     =   41675
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
      TabIndex        =   13
      Top             =   1560
      Width           =   1155
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
      TabIndex        =   12
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Client Name"
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
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Price"
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
      Top             =   3480
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Client Address"
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
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Print Purchase Request"
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
      TabIndex        =   8
      Top             =   240
      Width           =   2955
   End
End
Attribute VB_Name = "frmLogisticsPRAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPurchaseAdd_Click()
    query = "INSERT INTO tbl_tp_logistics(ID, purchase_id, quotation_id, purchase_date, equipment) VALUES ('','" & Trim(Me.txtPurchaseNumber) & "','" & Me.txtQuotationNumber & "','" & Me.dtpPurchaseDate & "', '" & Trim(Me.txtPurchaseEq) & "')"
    
    Connect.Execute (query)
    
    MsgBox "Successfully Added", vbInformation, frmLogisticsPRAdd.Caption
    
    Unload Me
End Sub
