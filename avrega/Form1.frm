VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Result"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Entet 3 value"
      Height          =   615
      Left            =   840
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Enter 2 value"
      Height          =   735
      Left            =   840
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter 1st value"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = (a + b + c) / 3
Text4.Text = Val(d)

End Sub
