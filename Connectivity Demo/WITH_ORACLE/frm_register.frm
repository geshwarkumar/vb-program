VERSION 5.00
Begin VB.Form frm_register 
   Caption         =   "User Registration..."
   ClientHeight    =   3420
   ClientLeft      =   7515
   ClientTop       =   4305
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Bell MT"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_register.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6090
   Begin VB.CommandButton btn_Delete 
      Caption         =   "Delete"
      Height          =   450
      Left            =   3840
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton btn_Update 
      Caption         =   "Update"
      Height          =   450
      Left            =   2160
      TabIndex        =   4
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txt_conpass 
      Height          =   450
      Left            =   3360
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txt_username 
      Height          =   450
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txt_password 
      Height          =   450
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton btn_submit 
      Caption         =   "Submit"
      Height          =   450
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lbl_confirmPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label lbl_title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "User Registration Demo"
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
      Left            =   -240
      TabIndex        =   8
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lbl_username 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1695
   End
End
Attribute VB_Name = "frm_register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
