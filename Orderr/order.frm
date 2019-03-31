VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "Bleeding Cowboys"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000FF&
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Enter Array Element"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "Assending Order"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Desending Order"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   8040
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, k As Double
Dim a(6) As Double
Private Sub Command1_Click()

For i = 0 To 5
   
    a(i) = InputBox("Enter 6 Element", "Array Element")
    Text1 = Text1 & a(i) & " " & vbNewLine
  
Next
End Sub

Private Sub Command2_Click()
For i = 0 To 5
    For j = 0 To 5
    If (a(j) > a(j + 1)) Then
    k = a(j)
    a(j) = a(j + 1)
    a(j + 1) = k
    End If
    Next
Next


For i = 1 To 6
    Text2 = Text2 & a(i) & vbNewLine
Next
End Sub

Private Sub Command3_Click()
For i = 0 To 5
    For j = 0 To 5
    If (a(j) < a(j + 1)) Then
    t = a(j)
    a(j) = a(j + 1)
    a(j + 1) = t
    End If
    Next
Next


For i = 0 To 5
    Text3 = Text3 & a(i) & vbNewLine
Next
End Sub

Private Sub Command4_Click()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
