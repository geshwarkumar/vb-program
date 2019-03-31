VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Student Records..."
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
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
   ScaleHeight     =   6555
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   9375
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "ADO.frx":0000
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   9375
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   8040
         Top             =   2280
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
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
         Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\23 ADO\Student.accdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\23 ADO\Student.accdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Student"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   14
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         DataField       =   "Mobile Number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         DataField       =   "Address"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6120
         TabIndex        =   10
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         DataField       =   "Name"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   6120
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         DataField       =   "Roll Number"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Mobile number"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Class name"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Student name"
         Height          =   375
         Left            =   4440
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Roll number"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Program to Acccess database using ADO & display the record in DataGrid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
On Error Resume Next
Adodc1.Recordset.Save
End Sub

Private Sub Command3_Click()
On Error Resume Next
Adodc1.Recordset.Update
End Sub

Private Sub Command4_Click()
On Error Resume Next
Adodc1.Recordset.Delete
End Sub

Private Sub Command5_Click()
Unload Me
End Sub
