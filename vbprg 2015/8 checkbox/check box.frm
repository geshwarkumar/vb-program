VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Using checkbox..."
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check7"
      Height          =   480
      Left            =   2280
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Font color"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Descrease font size"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Increase font sisze"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Enter a Text"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "         Using Check box"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click() 'for bold
If Check1.Value = 1 Then
Text1.FontBold = True
ElseIf Check1.Value = 0 Then
Text1.FontBold = False
End If
End Sub

Private Sub Check2_Click() 'for italic
If Check2.Value = 1 Then
Text1.FontItalic = True
ElseIf Check2.Value = 0 Then
Text1.FontItalic = False
End If
End Sub

Private Sub Check3_Click() 'for underline
If Check3.Value = 1 Then
Text1.FontUnderline = True
ElseIf Check3.Value = 0 Then
Text1.FontUnderline = False
End If
End Sub

Private Sub Check4_Click() 'for increace font size
If Check4.Value = 1 Then
Text1.FontSize = 18
ElseIf Check4.Value = 0 Then
Text1.FontSize = 12
End If
End Sub

Private Sub Check5_Click() 'for decrease font size
If Check5.Value = 1 Then
Text1.FontSize = 8
ElseIf Check5.Value = 0 Then
Text1.FontSize = 12
End If
End Sub

Private Sub Check6_Click() 'for font color
If Check6.Value = 1 Then
Text1.ForeColor = vbBlue
ElseIf Check6.Value = 0 Then
Text1.ForeColor = vbBlack
End If
End Sub


Private Sub Command2_Click() 'for exit
Unload Me
End Sub
