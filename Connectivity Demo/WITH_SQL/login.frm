VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_login 
   Caption         =   "Login Form..."
   ClientHeight    =   2715
   ClientLeft      =   7305
   ClientTop       =   4515
   ClientWidth     =   6495
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
   ScaleHeight     =   2715
   ScaleWidth      =   6495
   Begin VB.TextBox txt_password 
      Height          =   450
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txt_username 
      Height          =   450
      Left            =   1800
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   600
      Top             =   2280
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
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
      Left            =   4680
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton btn_submit 
      Caption         =   "Login"
      Height          =   450
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lbl_username 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   375
      Left            =   240
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btn_submit_Click()
Adodc1.Refresh

If Adodc1.Recordset.Fields("username").Value = txt_username.Text And Adodc1.Recordset.Fields("password").Value = txt_password.Text Then
    Me.Hide
    MsgBox "Succeful login..."
Else
    MsgBox "Unsucceed..."
    txt_username.SetFocus
End If
End Sub
