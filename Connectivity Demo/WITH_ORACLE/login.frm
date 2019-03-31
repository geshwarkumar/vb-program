VERSION 5.00
Begin VB.Form frm_login 
   Caption         =   "Login Form..."
   ClientHeight    =   2280
   ClientLeft      =   7305
   ClientTop       =   4515
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Bell MT"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6450
   Begin VB.CommandButton btn_cancel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   4680
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton btn_submit 
      Caption         =   "Login"
      Height          =   450
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txt_password 
      Height          =   450
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txt_username 
      Height          =   450
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl_username 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lbl_title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login Demo"
      BeginProperty Font 
         Name            =   "Bell MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
