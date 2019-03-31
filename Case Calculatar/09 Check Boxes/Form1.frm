VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   1350
   ClientTop       =   2310
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Caption         =   "Font Effects"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.CommandButton Command6 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Caption         =   "Change ForeColor"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   5400
         TabIndex        =   7
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option4 
            Caption         =   "RED"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   555
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "BLUE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Left            =   360
            TabIndex        =   9
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton Option6 
            Caption         =   "GREEN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   555
            Left            =   360
            TabIndex        =   8
            Top             =   1440
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Change Font Color"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   4935
         Begin VB.CheckBox Check1 
            Caption         =   "BOLD"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Italic"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1680
            TabIndex        =   5
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Underline"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3000
            TabIndex        =   4
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option7 
            Caption         =   "size increse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   240
            TabIndex        =   3
            Top             =   1200
            Width           =   1455
         End
         Begin VB.OptionButton Option8 
            Caption         =   "size decrise"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1680
            TabIndex        =   2
            Top             =   1200
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Vandemataram College Dhamtari"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   855
         TabIndex        =   13
         Top             =   600
         Width           =   6135
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   7800
         Y1              =   1680
         Y2              =   1680
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Label1.Font.Bold = True
End Sub

Private Sub Check2_Click()
Label1.Font.Italic = True
End Sub

Private Sub Check3_Click()
Label1.Font.Underline = True
End Sub

Private Sub Command1_Click()
Label1.Font.Bold = True
End Sub

Private Sub Command2_Click()
Label1.Font.Italic = True
End Sub

Private Sub Command3_Click()
Label1.Font.Underline = True
End Sub

Private Sub Command4_Click()
Label1.Font.Size = 30
End Sub

Private Sub Command5_Click()
Label1.Font.Size = 15
End Sub

Private Sub Command6_Click()
Label1.Font.Bold = False
Label1.Font.Italic = False
Label1.Font.Underline = False
Label1.Font.Size = 20
Label1.Font.Size = 20
Label1.ForeColor = vbBlack
Label1.BackColor = &H8000000F
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Option1_Click()
Label1.ForeColor = vbRed
End Sub

Private Sub Option2_Click()
Label1.ForeColor = vbBlue
End Sub

Private Sub Option3_Click()
Label1.ForeColor = vbGreen
End Sub

Private Sub Option4_Click()
Label1.ForeColor = vbRed
End Sub

Private Sub Option5_Click()
Label1.ForeColor = vbBlue
End Sub

Private Sub Option6_Click()
Label1.ForeColor = vbGreen
End Sub

Private Sub Option7_Click()
Label1.Font.Size = 30
End Sub

Private Sub Option8_Click()
Label1.Font.Size = 15
End Sub
