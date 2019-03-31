VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Functionality of line & shape control..."
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7365
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
   ScaleHeight     =   4365
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      ItemData        =   "ShapLine.frx":0000
      Left            =   4320
      List            =   "ShapLine.frx":0002
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Color"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   840
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Demonstrate the functionality of Line and Shape control using Common Dialog Control"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
CommonDialog1.ShowColor
Shape1.FillColor = CommonDialog1.Color
Shape1.BorderColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
List1.AddItem "Rectangle"
List1.AddItem "Square"
List1.AddItem "Oval"
List1.AddItem "Circle"
List1.AddItem "Round Rectangle"
List1.AddItem "Round Square"
End Sub

Private Sub List1_Click()
Select Case List1.ListIndex
    Case 0
        Shape1.Shape = 0
    Case 1
        Shape1.Shape = 1
    Case 2
        Shape1.Shape = 2
    Case 3
        Shape1.Shape = 3
    Case 4
        Shape1.Shape = 4
    Case 5
        Shape1.Shape = 5
End Select
End Sub

