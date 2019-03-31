VERSION 5.00
Begin VB.Form ChangeColorShape 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7455
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
   ScaleHeight     =   7185
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Shapes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
      Begin VB.OptionButton optShape 
         Caption         =   "ROUNDED SQUARE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   2295
      End
      Begin VB.OptionButton optShape 
         Caption         =   "ROUNDED RECTANGLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   2535
      End
      Begin VB.OptionButton optShape 
         Caption         =   "CIRCLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton optShape 
         Caption         =   "OVEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.OptionButton optShape 
         Caption         =   "SQUARE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton optShape 
         Caption         =   "RECTANGLE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Border Color"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   1
      Top             =   3360
      Width           =   1935
      Begin VB.OptionButton optColor 
         Caption         =   "GREEN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optColor 
         Caption         =   "BLUE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optColor 
         Caption         =   "RED"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Change Border color and Shape which selected into the list"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   2040
      Top             =   1320
      Width           =   2655
   End
End
Attribute VB_Name = "ChangeColorShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnExit_Click()
Unload Me
End Sub

Private Sub optColor_Click(Index As Integer)
If optColor(0).Value = True Then
    Shape1.FillStyle = 0
    Shape1.FillColor = vbRed
ElseIf optColor(1).Value = True Then
    Shape1.FillStyle = 0
    Shape1.FillColor = vbBlue
ElseIf optColor(2).Value = True Then
    Shape1.FillStyle = 0
    Shape1.FillColor = vbGreen
Else
    Shape1.FillStyle = vbTransparent
End If
End Sub

Private Sub optShape_Click(Index As Integer)
Shape1.Shape = Index
End Sub
