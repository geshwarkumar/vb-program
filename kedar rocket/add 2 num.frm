VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "add 2 num.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "OUTPUT SECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   10320
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
      Begin VB.TextBox Text4 
         BackColor       =   &H00808080&
         Height          =   1215
         Left            =   480
         TabIndex        =   13
         Top             =   5040
         Width           =   3375
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FF80FF&
         Caption         =   "/3"
         Height          =   1215
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "*"
         Height          =   1215
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "-"
         Height          =   1095
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "+"
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   4440
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   4440
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "INPUT SECTION"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   5760
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox Text3 
         Height          =   660
         Left            =   1440
         TabIndex        =   8
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   660
         Left            =   1320
         TabIndex        =   7
         Top             =   3600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   660
         Left            =   1320
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enter 3rd number"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   4440
         Width           =   4095
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enter 2nd number"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Enter 1st number"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   3960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3960
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FF00&
      Caption         =   "PROGRAM FOR ADD TWO NUMBER"
      Height          =   735
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   8535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = a + b + c
Text4.Text = Val(d)
End Sub

Private Sub Command2_Click()

Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = a - b - c
Text4.Text = Val(d)
End Sub

Private Sub Command3_Click()

Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = a * b * c
Text4.Text = Val(d)
End Sub

Private Sub Command4_Click()

Dim a, b, c, d As Integer
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = a + b + c / 3
Text4.Text = Val(d)
End Sub
