VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5310
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7590
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuMatrix_Addition 
      Caption         =   "Matrix_Addition"
   End
   Begin VB.Menu mnuMatrix_Subtraction 
      Caption         =   "Matrix_Subtraction"
   End
   Begin VB.Menu mnuMatrix_Multiplication 
      Caption         =   "Matrix_Multiplication"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuMatrix_Addition_Click()
Form1.Show
Form2.Hide
Form3.Hide
End Sub

Private Sub mnuMatrix_Multiplication_Click()
Form3.Show
Form1.Hide
Form2.Hide
End Sub

Private Sub mnuMatrix_Subtraction_Click()
Form2.Show
Form1.Hide
Form3.Hide
End Sub
