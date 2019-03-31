VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Change back color..."
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Left            =   4800
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      Left            =   480
      Max             =   300
      TabIndex        =   3
      Top             =   3600
      Width           =   4935
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   480
      Max             =   300
      TabIndex        =   2
      Top             =   2880
      Width           =   4935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   480
      Max             =   300
      TabIndex        =   1
      Top             =   2160
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Change back color of any controls ex: lable"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Lable1"
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub HScroll1_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Label1.BackColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
