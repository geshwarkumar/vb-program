VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_changepassword 
   Caption         =   "Change Password"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Bell MT"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_changepassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3960
      Top             =   3360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=system;Data Source=localhost;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=system;Data Source=localhost;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "system"
      Password        =   "admin"
      RecordSource    =   "select * from userdetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bell MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton btn_cancel 
      Caption         =   "Cancel"
      Height          =   450
      Left            =   2280
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt_conpass 
      Height          =   450
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txt_newpass 
      DataSource      =   "Adodc1"
      Height          =   450
      Left            =   3600
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txt_oldpass 
      DataSource      =   "Adodc1"
      Height          =   450
      Left            =   3600
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton btn_change 
      Caption         =   "Submit"
      Height          =   450
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txt_username 
      DataSource      =   "Adodc1"
      Height          =   450
      Left            =   3600
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label lbl_conpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lbl_newpass 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lbl_oldpass 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lbl_username 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lbl_title 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password Demo"
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
      Left            =   -120
      TabIndex        =   6
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_changepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
