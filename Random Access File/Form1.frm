VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   855
      Left            =   8040
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read"
      Height          =   855
      Left            =   4920
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim P As person
    Open "d:\hello.doc" For Random As #1 Len = Len(P)

    P.LName = "Ram"
    P.FName = "Dashrath"
    P.Age = 9
    Put 1, , P

    P.LName = "Ravan"
    P.FName = "YoYo"
    P.Age = 4
    Put 1, , P

    Close #1
End Sub

Private Sub Command2_Click()
Dim P As person
    Open "d:\hello.doc" For Random As 1 Len = Len(P)
    For i = 1 To Int(LOF(1) / Len(P))
    Get 1, i, P
    Print P.LName, P.LName, P.Age = 5
    Next
    Close 1
End Sub

Private Sub Command3_Click()
Dim P As person
    Open "d:\hello.doc" For Random As 1 Len = Len(P)
    Get 1, 2, P
    P.Age = 5
    Put 1, 2, P
    Close 1
End Sub
