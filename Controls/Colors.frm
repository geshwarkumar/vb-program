VERSION 5.00
Begin VB.Form Colors 
   Caption         =   "Color change"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8325
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
   ScaleHeight     =   6420
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll3 
      Height          =   495
      Left            =   2160
      Max             =   1000
      TabIndex        =   2
      Top             =   4920
      Width           =   4455
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   495
      Left            =   2160
      Max             =   1000
      TabIndex        =   1
      Top             =   4200
      Width           =   4455
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   2160
      Max             =   1000
      TabIndex        =   0
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000007&
      FillStyle       =   0  'Solid
      Height          =   2415
      Left            =   2520
      Shape           =   2  'Oval
      Top             =   600
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Green"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Blue"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Red"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Colors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Shape1.FillColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll1_Change()
Shape1.FillColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll2_Change()
Shape1.FillColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub

Private Sub HScroll3_Change()
Shape1.FillColor = RGB(HScroll1.Value, HScroll2.Value, HScroll3.Value)
End Sub
