VERSION 5.00
Begin VB.Form Interest 
   Caption         =   "Calculate Interest"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8235
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5610
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCi 
      Caption         =   "Compound Interest"
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   2400
      Width           =   2655
   End
   Begin VB.OptionButton optSi 
      Caption         =   "Simple Interest"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   3000
      TabIndex        =   10
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txttime 
      Height          =   435
      Left            =   2880
      TabIndex        =   6
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox txtrate 
      Height          =   435
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtpr 
      Height          =   435
      Left            =   2880
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Total value of Principal and Interest"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Time"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Rate"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Principal Amount"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Program to Calculate Interest"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Interest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim p, r, t As Integer, i As Double
p = Val(txtpr.Text)
r = Val(txtrate.Text)
t = Val(txttime.Text)

If (optSi = True) Then

    i = p * r * t / 100
    txttotal.Text = Val(i) + p
    
Else
    
   If (optCi = True) Then
        i = p * ((1 + 0.01 * r) ^ t)
        txttotal.Text = Val(i)
   End If

End If

End Sub

Private Sub Command2_Click()
End
End Sub
