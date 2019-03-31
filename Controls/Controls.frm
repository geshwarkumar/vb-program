VERSION 5.00
Begin VB.Form Controls 
   Caption         =   "ListBox and ComboBox Controls..."
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
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
   ScaleHeight     =   6810
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      ItemData        =   "Controls.frx":0000
      Left            =   840
      List            =   "Controls.frx":0002
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Multiplication Table"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Program for demonstrate the use of listbox and combobox"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
'Dim i As Integer
List1.Clear
For i = 1 To 12
    List1.AddItem (Combo1.ListIndex + 1) & " * " & i & " = " & i * (Combo1.ListIndex + 1)
Next i
End Sub

Private Sub Form_Load()
Dim i As Integer
Combo1.Clear
For i = 1 To 100
    Combo1.AddItem "Table of " & i
Next i
End Sub
