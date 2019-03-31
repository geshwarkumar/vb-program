VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Matrix Multiply"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Enter First Matrix Element"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Enter Second Matrix Element"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   4560
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(2, 2), b(2, 2), c(2, 2) As Integer
Dim i, j, k As Integer
Private Sub Command1_Click()
For i = 0 To 2
For j = 0 To 2
a(i, j) = InputBox("enter 1st matrix element")
Next j, i

For i = 0 To 2
For j = 0 To 2
Text1 = Text1 & a(i, j) & vbTab
Next j
Text1 = Text1 & vbNewLine
Next i
End Sub

Private Sub Command2_Click()
For i = 0 To 2
For j = 0 To 2
b(i, j) = InputBox("enter 2st matrix element")
Next j, i

For i = 0 To 2
For j = 0 To 2
Text2 = Text2 & b(i, j) & vbTab
Next j
Text2 = Text2 & vbNewLine
Next i
End Sub

Private Sub Command3_Click()
For i = 0 To 2
For j = 0 To 2
c(i, j) = 0
For k = 0 To 2
c(i, j) = c(i, j) + a(i, k) * b(k, j)
Next k, j, i

For i = 0 To 2
For j = 0 To 2
Text3 = Text3 & c(i, j) & vbTab
Next j
Text3 = Text3 & vbNewLine
Next i
End Sub

