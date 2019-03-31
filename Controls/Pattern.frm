VERSION 5.00
Begin VB.Form Pattern 
   Caption         =   "Form1"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5865
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
   ScaleHeight     =   5160
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4320
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Pyramid of stars"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "No of line"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim a As Integer
a = Val(Text1.Text)

For i = 0 To a

    For j = a To i
        Print " ";
    Next
    
    For k = 1 To a
        Print "*";
    Next
    Print " "
    
Next
    
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


