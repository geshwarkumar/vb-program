VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   3060
   ClientTop       =   3225
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5985
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, n As Integer
a = 0
n = 1
While (n <= 10)
a = a + 2
'Print a;
Print a
MsgBox a
Text1 = Text1.Text & a
Label1.Caption = Label1.Caption & a
n = n + 1
Wend
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
