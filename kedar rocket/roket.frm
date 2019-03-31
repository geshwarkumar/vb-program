VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   13170
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   3240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "bottom"
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "top"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "start"
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "stop"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1920
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   120
      Picture         =   "roket.frx":0000
      Top             =   600
      Width           =   1500
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
End Sub



Private Sub Command4_Click()
If (Timer1.Enabled = False) Then

Timer2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Image1.Left = Image1.Left + 100
End Sub

Private Sub Timer2_Timer()
Image1.Top = Image1.Top + 100
End Sub
