VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   13320
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF80FF&
      Caption         =   "Frame1"
      Height          =   7575
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      Begin VB.CommandButton Command2 
         BackColor       =   &H008080FF&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "click"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6000
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   3840
         Width           =   5535
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5880
         TabIndex        =   3
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Enter any terms for         SERIES"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         TabIndex        =   2
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "     Program for fibbonaccy series"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   855
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   7335
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i, n, s, l, f As Integer
n = Val(Text1.Text)
f = 0
s = 1
For i = 1 To n
l = f + s
Text2.Text = Text2.Text & " " & (f)
f = s
s = l
Next

End Sub


Private Sub Command2_Click()
Text1 = ""
Text2 = ""
End Sub

