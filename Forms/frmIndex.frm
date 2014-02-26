VERSION 5.00
Begin VB.Form Index 
   Caption         =   "Index"
   ClientHeight    =   8280
   ClientLeft      =   2115
   ClientTop       =   2100
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   12765
   Begin VB.Menu mnuAccounting 
      Caption         =   "&Accounting"
   End
   Begin VB.Menu mnuEngineering 
      Caption         =   "&Engineering"
   End
   Begin VB.Menu mnuLogistics 
      Caption         =   "&Logistics"
   End
   Begin VB.Menu mnuClients 
      Caption         =   "&Clients"
   End
   Begin VB.Menu mnuSales 
      Caption         =   "&Sales"
   End
End
Attribute VB_Name = "Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub mnuAccounting_Click()
    IndexAccounting.Show
End Sub

Private Sub mnuClients_Click()
    Clients.Show
End Sub

Private Sub mnuEngineering_Click()
    IndexEngineering.Show
End Sub

Private Sub mnuLogistics_Click()
    IndexLogistics.Show
End Sub

Private Sub mnuSales_Click()
    IndexSales.Show
End Sub
