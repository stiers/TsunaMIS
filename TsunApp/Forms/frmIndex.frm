VERSION 5.00
Begin VB.Form frmIndex 
   Caption         =   "Index"
   ClientHeight    =   8085
   ClientLeft      =   4935
   ClientTop       =   1995
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   6585
   Begin VB.Menu mnuUsers 
      Caption         =   "Users"
      Begin VB.Menu mnuUserAll 
         Caption         =   "All Users"
      End
      Begin VB.Menu mnuUserNew 
         Caption         =   "Add New"
      End
      Begin VB.Menu mnuUserProfile 
         Caption         =   "Your Profile"
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuUserAll_Click()
    frmUsers.Show
End Sub

Private Sub mnuUserNew_Click()
    frmUserNew.Show
End Sub

Private Sub mnuUserProfile_Click()
    frmProfile.Show
End Sub
