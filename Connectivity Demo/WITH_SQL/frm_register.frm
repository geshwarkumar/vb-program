VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_register 
   Caption         =   "User Registration..."
   ClientHeight    =   4110
   ClientLeft      =   7515
   ClientTop       =   4305
   ClientWidth     =   6615
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
   ScaleHeight     =   4110
   ScaleWidth      =   6615
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   450
      Left            =   4920
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc adoReg 
      Height          =   495
      Left            =   240
      Top             =   3480
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
   Begin VB.CommandButton btn_Delete 
      Caption         =   "Delete"
      Height          =   450
      Left            =   3360
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton btn_Update 
      Caption         =   "Update"
      Height          =   450
      Left            =   1680
      TabIndex        =   4
      Top             =   3000
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
      DataField       =   "USERNAME"
      DataSource      =   "adoReg"
      Height          =   450
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txt_password 
      DataField       =   "PASSWORD"
      DataSource      =   "adoReg"
      Height          =   450
      Left            =   3360
      TabIndex        =   1
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton btn_submit 
      Caption         =   "Submit"
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   3000
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
Dim con As ADODB.Connection
Dim res As ADODB.Recordset

Private Sub btn_Delete_Click()
On Error Resume Next
adoReg.Recordset.Delete
btn_Cleaar_Click
MsgBox "Record delete..."
End Sub

Private Sub btn_submit_Click()
On Error Resume Next
adoReg.Recordset.AddNew
MsgBox "Save succeed..."
End Sub

Private Sub btn_Update_Click()
On Error Resume Next
adoReg.Recordset.Update
MsgBox "Update succeed..."
End Sub

Private Sub btnClear_Click()
txt_username = ""
txt_password = ""
txt_conpass = ""
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    con.Open "Provider=MSDAORA.1;Password=admin;User ID=system;Data Source=localhost;Persist Security Info=False"
    con.CursorLocation = adUseClient
    MsgBox "Connection stablished"
End Sub

