VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   975
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click To Color Change"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6375
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   855
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "WAP for change Hsroll Bar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   2865
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.BackColor = HScroll1.Value
Label1.BackColor = HScroll1.Value
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.BackColor = 0
Label1.BackColor = 0
End Sub

Private Sub HScroll1_Change()
HScroll1.Min = 0
HScroll1.Max = 255
End Sub
