VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Window"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   9405
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMDI 
      Caption         =   "&MDI Menu"
      Begin VB.Menu mnuopen 
         Caption         =   "MDI Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "MDI Exit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuopen_Click()
    Form1.Show
End Sub
