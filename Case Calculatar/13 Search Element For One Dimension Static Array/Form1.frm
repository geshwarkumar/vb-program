VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Search an Element For a One Dimension Static Array"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   4080
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Enter Five Element"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   450
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(5) As Integer
Dim i, d As Integer
Private Sub Command2_Click()
d = InputBox("Enter For Search")
Text2 = d
For i = 1 To 5
If (a(i) = d) Then
MsgBox "Position - " & i & vbNewLine & "Value - " & a(i)
Exit Sub
End If
Next
If (i > 5) Then
MsgBox "No Record"
End If

End Sub

Private Sub Command1_Click()
For i = 0 To 4
a(i) = InputBox("Enter Five Element")
Text1 = Text1 & a(i) & vbNewLine
Next
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
