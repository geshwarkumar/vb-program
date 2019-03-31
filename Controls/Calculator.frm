VERSION 5.00
Begin VB.Form Calculator 
   Caption         =   "Simple Calculator"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3135
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
   ScaleHeight     =   4260
   ScaleWidth      =   3135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2655
      Begin VB.CommandButton Command2 
         Caption         =   "Off"
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "AC"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "+"
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton btnDivid 
         Caption         =   "/"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton btnMinus 
         Caption         =   "-"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnMult 
         Caption         =   "*"
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton btnEual 
         Caption         =   "="
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton btnDot 
         Caption         =   "."
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   1320
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   1320
         TabIndex        =   9
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   8
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   6
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   5
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton btn 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   3
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtinput 
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Simple Calculator"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1, op2 As Double
Dim opr As String
Private Sub btn_Click(Index As Integer)
txtinput.Text = txtinput.Text + btn(Index).Caption

If cleardisplay Then
    txtinput.Text = ""
    cleardisplay = False
End If

End Sub

Private Sub btnAdd_Click()
op1 = Val(txtinput.Text)
opr = "+"
txtinput.Text = ""
End Sub

Private Sub btnDivid_Click()
op1 = Val(txtinput.Text)
opr = "/"
txtinput.Text = ""
End Sub

Private Sub btnDot_Click()

If InStr(txtinput.Text, ".") Then
    Exit Sub
Else
    txtinput.Text = txtinput.Text + "."
End If

End Sub

Private Sub btnEual_Click()

op2 = Val(txtinput.Text)
Select Case opr
    Case "+": txtinput.Text = op1 + op2
    Case "-": txtinput.Text = op1 - op2
    Case "*": txtinput.Text = op1 * op2
    Case "/": txtinput.Text = op1 / op2
End Select

End Sub

Private Sub btnMinus_Click()
op1 = Val(txtinput.Text)
opr = "-"
txtinput.Text = ""
End Sub

Private Sub btnMult_Click()
op1 = Val(txtinput.Text)
opr = "*"
txtinput.Text = ""
End Sub

Private Sub Command1_Click()
txtinput.Text = ""
End Sub

Private Sub Command2_Click()
txtinput.Text = txtinput.Text - btn(Index).Caption

If cleardisplay = False Then
    txtinput.Text = btn(Index).Caption
    'cleardisplay = False
End If
'Unload Me
End Sub

Private Sub Form_Load()
txtinput.Text = ""
End Sub
