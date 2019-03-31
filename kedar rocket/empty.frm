VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16020
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   16020
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   2400
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "OUT PUT"
      Height          =   8295
      Left            =   12480
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   6255
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Caption         =   "Descending"
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   " OUT PUT"
      Height          =   8295
      Left            =   3840
      TabIndex        =   1
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   6375
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
         Caption         =   "Ascending"
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Height          =   9015
      Left            =   7920
      TabIndex        =   0
      Top             =   840
      Width           =   4335
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF8080&
         Caption         =   "-->"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8040
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00008000&
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   8160
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "<--"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8040
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   5400
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Click"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H0080C0FF&
         Caption         =   "Click in block for out put"
         Height          =   1215
         Left            =   360
         TabIndex        =   12
         Top             =   6480
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Given input is--"
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   4440
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   " Click for Enter       Number"
         Height          =   1095
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF00FF&
         Caption         =   "Program for order"
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   3855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, t As Double
Dim a(6) As Double

Private Sub Command1_Click()
For i = 0 To 5
   
    a(i) = InputBox("Enter 6 Element")
    Text1 = Text1 & a(i) & " "
  
Next
End Sub

Private Sub Command2_Click()
Text2 = ""
Text3 = ""
For i = 0 To 5
    For j = 0 To 5
    If (a(j) > a(j + 1)) Then
    t = a(j)
    a(j) = a(j + 1)
    a(j + 1) = t
    End If
    Next
Next


For i = 1 To 6
    Text2 = Text2 & a(i) & vbNewLine
Next
Text3 = ""
End Sub

Private Sub Command3_Click()
Text2 = ""
Text3 = ""
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
For i = 0 To 5
    For j = 0 To 5
    If (a(j) > a(j + 1)) Then
    t = a(j)
    a(j) = a(j + 1)
    a(j + 1) = t
    End If
    Next
Next


For i = 1 To 6
    Text2 = Text2 & a(i) & vbNewLine
Next
End Sub

Private Sub Command4_Click()
Text2 = ""
Text3 = ""
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
Text2 = ""
End Sub

Private Sub Timer1_Timer()
If (Form1.BackColor = vbBlack) Then
Form1.BackColor = vbBlue

 ElseIf (Form1.BackColor = vbBlue) Then
Form1.BackColor = vbYellow

ElseIf (Form1.BackColor = vbYellow) Then
Form1.BackColor = vbGreen

ElseIf (Form1.BackColor = vbGreen) Then
Form1.BackColor = vbCyan

ElseIf (Form1.BackColor = vbBlue) Then
Form1.BackColor = vbMagenta

Else

Form1.BackColor = vbBlack
End If


End Sub
