VERSION 5.00
Begin VB.Form ListBox 
   Caption         =   "List Box"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton btnRemove 
      Caption         =   "<"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton btnRemoveAll 
      Caption         =   "<<"
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton btnAddAll 
      Caption         =   ">>"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   ">"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   3525
      Left            =   4200
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   3525
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
Dim a As Integer
i = 0
If List1.SelCount <> 0 Then
    Do While List1.SelCount > 0
        If List1.Selected(i) Then
            List2.AddItem List1.List(i)
            List1.RemoveItem i
        Else
            i = i + 1
        End If
    Loop
Else
    MsgBox "First item selected"
End If
End Sub

Private Sub btnAddAll_Click()
If List1.ListCount = 0 Then
MsgBox "No data found"
Else
For i = 0 To List1.ListCount - 1
    List2.AddItem List1.List(i)
Next
    List1.Clear
End If
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub btnRemove_Click()
Dim i As Integer
i = 0
If List2.SelCount <> 0 Then
    Do While List2.SelCount > 0
        If List2.Selected(i) Then
            List1.AddItem List2.List(i)
            List2.RemoveItem i
        Else
            i = i + 1
        End If
    Loop
Else
    MsgBox "First item selected"
End If
End Sub

Private Sub btnRemoveAll_Click()
If List2.ListCount = 0 Then
MsgBox "No data found"
For i = 0 To List2.ListCount - 1
    List1.AddItem List2.List(i)
Next
    List2.Clear
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
List1.Clear
For i = 0 To 15
List1.AddItem "item" & i
Next
End Sub
