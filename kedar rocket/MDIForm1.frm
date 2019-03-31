VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu nmul 
      Caption         =   "Loop"
      Begin VB.Menu mnuf 
         Caption         =   "fibonecy"
      End
      Begin VB.Menu mnuo 
         Caption         =   "order"
      End
      Begin VB.Menu mnu 
         Caption         =   "rocket"
      End
   End
   Begin VB.Menu add 
      Caption         =   "add"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub add_Click()
Form4.Show
End Sub

Private Sub mnu_Click()
Form3.Show
End Sub

Private Sub mnuf_Click()
Form2.Show
End Sub

Private Sub mnuo_Click()
Form1.Show
End Sub
