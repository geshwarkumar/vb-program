VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Connectivity 
   Caption         =   "Database connectivity"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10050
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
   ScaleHeight     =   7170
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Connectivity.frx":0000
      Height          =   2535
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Frame Frame2 
      Caption         =   "Action area"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5040
      TabIndex        =   19
      Top             =   1320
      Width           =   1575
      Begin VB.CommandButton btnReport 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton btnPrev 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnNext 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton btnDel 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add New"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Working area"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   4695
      Begin VB.TextBox txtSem 
         DataField       =   "Semester"
         DataSource      =   "Adodc1"
         Height          =   435
         Left            =   2520
         TabIndex        =   5
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtCourse 
         DataField       =   "Course"
         DataSource      =   "Adodc1"
         Height          =   435
         Left            =   2520
         TabIndex        =   4
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtFname 
         DataField       =   "FName"
         DataSource      =   "Adodc1"
         Height          =   435
         Left            =   2520
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         DataField       =   "SName"
         DataSource      =   "Adodc1"
         Height          =   435
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtRollno 
         DataField       =   "Rollno"
         DataSource      =   "Adodc1"
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Semester"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Course"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Father`s name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Student name"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Roll number"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5160
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      BackColor       =   -2147483638
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\MCA-1\VB\Demo.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\MCA-1\VB\Demo.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Student"
      Caption         =   ""
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Student Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   18
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Example of connectivity using ado control"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "Connectivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
On Error Resume Next
Adodc1.Recordset.AddNew
End Sub

Private Sub btnDel_Click()
On Error Resume Next
Adodc1.Recordset.Delete
MsgBox "Record delete in databse..."
Adodc1.Refresh
End Sub

Private Sub btnNext_Click()
On Error Resume Next
If Not Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MovePrevious
End If
End If
End Sub

Private Sub btnPrev_Click()
On Error Resume Next
If Not Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then
    Adodc1.Recordset.MoveNext
End If
End If
End Sub

Private Sub btnReport_Click()
DataReport1.Show
Connectivity.Visible = False
End Sub

Private Sub btnSave_Click()
On Error Resume Next
Adodc1.Recordset.Save
MsgBox "Record save succeed..."
End Sub

Private Sub btnUpdate_Click()
On Error Resume Next
Adodc1.Recordset.Update
MsgBox "Update successful..."
Adodc1.Refresh
End Sub
