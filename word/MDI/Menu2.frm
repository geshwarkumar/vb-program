VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MDI Child Window"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   Begin VB.Menu mnuChild 
      Caption         =   "Child Menu"
      Begin VB.Menu mnucopen 
         Caption         =   "Child Open"
      End
      Begin VB.Menu mnuCSave 
         Caption         =   "Child Save"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Child Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClose_Click()
Unload Me
End Sub
