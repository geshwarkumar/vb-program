VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sorting element in ascending order..."
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
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
   ScaleHeight     =   4290
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   3240
      TabIndex        =   9
      Top             =   2760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Soting List"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input List"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   4680
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "How many element in list(max 100)"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Sorted element are:"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Given element are:"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sorting the given element of one dimentional array in ascending order"
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
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(100), i, j, temp, num As Integer
Private Sub Command1_Click()
num = Val(Text1.Text)
For i = 1 To num Step 1
    a(i) = InputBox("Enter element for Sorting")
Next i
End Sub

Private Sub Command2_Click()
For i = 1 To num Step 1
    For j = 1 To num - 1 Step 1
        If a(j) > a(j + 1) Then
            temp = a(j)
            a(j) = a(j + 1)
            a(j + 1) = temp
        End If
    Next j
Next i
For i = 1 To num Step 1
    Text2.Text = Text2.Text & " " & a(i)
    Text3.Text = Text3.Text & " " & a(i)
Next i
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
