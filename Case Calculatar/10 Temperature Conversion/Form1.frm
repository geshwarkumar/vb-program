VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Temperature Conversion"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "C to F"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "F to C"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   4
         Top             =   1725
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   3720
         TabIndex        =   0
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
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
         TabIndex        =   2
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Enter Temperature"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   135
         TabIndex        =   3
         Top             =   840
         Width           =   2745
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c, f As Double

Private Sub Command1_Click()
Text1 = ""
Text2 = ""
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Option1_Click()
c = Text1
f = ((9 / 5) * c) + 32
Text2 = f
End Sub

Private Sub Option2_Click()
f = Text1
c = 5 / 9 * (f - 32)
Text2 = c
End Sub
