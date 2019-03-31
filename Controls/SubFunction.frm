VERSION 5.00
Begin VB.Form SubFunction 
   Caption         =   "Subroutine and function"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10125
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
   ScaleHeight     =   9660
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton btnResult 
      Caption         =   "call"
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "SubFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub btnResult_Click()
Dim a As Integer
a = 1
Print "Value of a before callbyvalue is " & a
Call callbyvalue(a)
Print "Value of a after callbyvalue is " & a
Print

Print
Print "Value of a before callbyvalue1 is " & a
Call callbyvalue1(a)
Print "Value of a after callbyvalue1 is " & a
Print

Print "Value of a before callbyref is " & a
Call callbyref(a)
Print "Value of a after callbyref is " & a
Print

Print
Print "Value of a before callbyref1 is " & a
Call callbyref1(a)
Print "Value of a after callbyref1 is " & a
Print

Print
Print "Value of a before callbyvalue1 is " & a
Call callbyvalue1(a)
Print "Value of a after callbyvalue1 is " & a
Print

Print
Print "Value of a before callbyref1 is " & a
Call callbyref1(a)
Print "Value of a after callbyref1 is " & a
Print
End Sub
Private Sub callbyvalue(ByVal x As Integer)
Print "Initial value x in callbyvalue is" & x
x = x * 3
Print "Last value of x in callbyvalue is" & x
End Sub
Private Sub callbyref(ByRef y As Integer)
Print "Initial value y in callbyref is" & y
y = y * 5
Print "Last value of y in callbyref is" & y
End Sub
Private Sub callbyvalue1(ByVal x1 As Integer)
Print "Initial value x1 in callbyvalue1 is" & x1
x1 = x1 * 3
Print "Last value of x1 in callbyvalue1 is" & x1
End Sub
Private Sub callbyref1(ByRef y1 As Integer)
Print "Initial value y1 in callbyref1 is" & y1
y1 = y1 * 5
Print "Last value of y1 in callbyref1 is" & y1
End Sub

