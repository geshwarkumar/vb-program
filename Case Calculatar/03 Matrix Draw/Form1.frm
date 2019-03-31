VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Matrix Format"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "Click For Matrix"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
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
         Height          =   615
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000FF&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   2175
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
         Height          =   615
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   3375
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n, i, b As Integer
n = Val(Text1)
For i = 1 To 10
b = n * i
Text2 = Text2 & vbNewLine & n & "*" & i & "=" & b
Next
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Form1.Cls
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
