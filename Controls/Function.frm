VERSION 5.00
Begin VB.Form Function 
   Caption         =   "Using Function..."
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4830
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
   ScaleHeight     =   3270
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   435
      Left            =   3000
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtNum 
      Height          =   435
      Left            =   3000
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton btnCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Result"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Enter the number"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Program to calculate factorial "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function fact(num As Integer) As Integer
    If num = 1 Then
        fact = 1
    Else
        fact = fact(num - 1) * num
    End If
End Function
Private Sub btnCalculate_Click()
txtResult.Text = Str(fact(Val(txtNum.Text)))
End Sub

Private Sub btnClear_Click()
txtNum.Text = ""
txtResult.Text = ""
End Sub

Private Sub btnExit_Click()
Unload Me
End Sub

