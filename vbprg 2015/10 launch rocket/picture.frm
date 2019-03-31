VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "Launch a rocket..."
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000009&
      Caption         =   "Exit"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6600
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Fly"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   5640
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      Picture         =   "picture.frx":0000
      ScaleHeight     =   2895
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Launch a rocket using picture box and timer control"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'for fly
Timer1.Enabled = True
End Sub

Private Sub Command2_Click() 'for stop
Timer1.Enabled = False
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Timer1_Timer() 'for picturebox
Picture1.Top = Picture1.Top - 10
End Sub
