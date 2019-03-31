VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   975
      Left            =   4920
      TabIndex        =   4
      Top             =   3600
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   855
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   840
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   4560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1215
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Text1.BackColor = HScroll1.Value
Label1.BackColor = HScroll1.Value
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.BackColor = 0
Label1.BackColor = 0
End Sub

Private Sub HScroll1_Change()
HScroll1.Min = 0
HScroll1.Max = 255
End Sub

