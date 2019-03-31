VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Open File"
      Height          =   855
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete File"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reading File"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim studentname As String
Dim intMsg As String
Private Sub Command1_Click()
'To read the file
Text1.Text = ""
Dim variable1 As String
On Error GoTo file_error
Open "D:\Liew Folder\sample.txt" For Input As #1
Do
Input #1, variable1
Text1.Text = Text1.Text & variable1 & vbCrLf
Loop While Not EOF(1)
Close #1

Exit Sub
file_error:
MsgBox (Err.Description)
End Sub
Private Sub Command2_Click()
'To delete the file
On Error GoTo delete_error
Kill "D:\Liew Folder\sample.txt"
Exit Sub
delete_error:
MsgBox (Err.Description)
End Sub
Private Sub Command3_Click()
End
End Sub
'Private Sub create_Click()
'To create the file or open the file for new data entry
'Open "D:\Liew Folder\sample.txt" For Append As #1
'intMsg = MsgBox("File sample.txt opened")
'Do
'studentname = InputBox("Enter the student Name or type finish to end")
'If studentname = "finish" Then
'Exit Do
'End If
'Write #1, studentname & vbCrLf
'intMsg = MsgBox("Writing " & studentname & " to sample.txt ")
'Loop
'Close #1
'intMsg = MsgBox("File sample.txt closed")
'End Sub

Private Sub Command4_Click()
'To create the file or open the file for new data entry
Open "D:\Liew Folder\sample.txt" For Append As #1
intMsg = MsgBox("File sample.txt opened")
Do
studentname = InputBox("Enter the student Name or type finish to end")
If studentname = "finish" Then
Exit Do
End If
Write #1, studentname & vbCrLf
intMsg = MsgBox("Writing " & studentname & " to sample.txt ")
Loop
Close #1
intMsg = MsgBox("File sample.txt closed")
End Sub

Private Sub Form_Load()

On Error GoTo Openfile_error
Open "D:\Liew Folder\sample.txt" For Input As #1
Close #1
Exit Sub
Openfile_error:
MsgBox (Err.Description), , "Please create a new file"
Create.Caption = "Create File"
End Sub

