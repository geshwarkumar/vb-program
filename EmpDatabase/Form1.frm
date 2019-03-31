VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee's Information System Version 1.0"
   ClientHeight    =   6615
   ClientLeft      =   165
   ClientTop       =   240
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&QUIT"
      Height          =   495
      Left            =   6120
      TabIndex        =   29
      ToolTipText     =   "Click here to quit the program."
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtdepartment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ABOUT"
      Height          =   495
      Left            =   6120
      TabIndex        =   22
      ToolTipText     =   "Click here to know the author of the program."
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Caption         =   "Employee's Information System Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6015
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtDob 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0000C000&
         Height          =   4575
         Left            =   5280
         TabIndex        =   10
         Top             =   240
         Width           =   2415
         Begin VB.CommandButton cmdSave 
            Caption         =   "&SAVE"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "Click here to save the records in the database."
            Top             =   240
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&SEARCH"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Click here to find the record in the database."
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "&UPDATE"
            Height          =   495
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Click here to update the record in the database."
            Top             =   1680
            Width           =   2175
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&DELETE"
            Height          =   495
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Click here to erase the record in the database."
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "ADD"
            Height          =   495
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0000C000&
         Height          =   855
         Left            =   240
         TabIndex        =   5
         Top             =   4920
         Width           =   7455
         Begin VB.CommandButton cmdLast 
            Caption         =   "Move Last"
            Height          =   495
            Left            =   5520
            TabIndex        =   9
            ToolTipText     =   "Click here to move to the last record."
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "Move Previous"
            Height          =   495
            Left            =   3720
            TabIndex        =   8
            ToolTipText     =   "Click here to move to the previous record."
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Move Next"
            Height          =   495
            Left            =   1920
            TabIndex        =   7
            ToolTipText     =   "Click here to move to the next record."
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdFirst 
            Caption         =   "Move First"
            Height          =   495
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Click here to move to the first record."
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   2760
         Width           =   2895
      End
      Begin VB.TextBox txtCity 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtNo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label10 
         BackColor       =   &H0000C000&
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Telephone Number "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   2880
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Date of Birth "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "City Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Employee Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         Caption         =   "Employee Number "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1740
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   27
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Telephone Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   26
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Telephone Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   25
      Top             =   3480
      Width           =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H0000C000&
      Caption         =   "Telephone Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   24
      Top             =   3480
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoconn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdAdd_Click()
    rs.AddNew
    txtNo.Text = ""
    txtName.Text = ""
    txtCity.Text = ""
    txtDob.Text = ""
    txtPhone.Text = ""
    txtdepartment.Text = ""
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdDelete.Enabled = False
    cmdSearch.Enabled = False
    cmdUpdate.Enabled = False
    cmdAdd.Visible = False
    cmdSave.Visible = True
End Sub

Private Sub cmdDelete_Click()
    Dim ans As String, str As String
    ans = MsgBox("Do you really want to delete the current record?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        adoconn.Execute ("delete from emp where e_no=" & txtNo.Text)
        MsgBox ("The record has been deleted successfully.")
        Set rs = Nothing
        str = "select * from emp"
        rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
        rs.MoveFirst
        txtNo.Text = rs(0)
        txtName.Text = rs(1)
        txtCity.Text = rs(2)
        txtDob.Text = rs(4)
        txtPhone.Text = rs(3)
        txtdepartment.Text = rs(5)
    End If
End Sub

Private Sub cmdFirst_Click()
    rs.MoveFirst
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
End Sub

Private Sub cmdLast_Click()
    rs.MoveLast
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
End Sub

Private Sub cmdNext_Click()
    rs.MoveNext
    If rs.EOF = True Then
        MsgBox "This is the last record.", vbExclamation, "Note it..."
        rs.MoveLast
    End If
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
End Sub

Private Sub cmdPrevious_Click()
    rs.MovePrevious
    If rs.BOF = True Then
        MsgBox "This is the first record.", vbExclamation, "Note it..."
        rs.MoveFirst
    End If
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
End Sub

Private Sub cmdSave_Click()
    rs(0) = txtNo.Text
    rs(1) = txtName.Text
    rs(2) = txtCity.Text
    rs(4) = txtDob.Text
    rs(3) = txtPhone.Text
    rs(5) = txtdepartment.Text
    rs.Update
    MsgBox "The record has been saved successfully.", , "ADD"
    cmdFirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    cmdDelete.Enabled = True
    cmdSearch.Enabled = True
    cmdUpdate.Enabled = True
    cmdSave.Visible = False
    cmdAdd.Visible = True
End Sub

Private Sub cmdSearch_Click()
    Dim key As Integer, str As String
    key = InputBox("Enter the Employee No whose details u want to know: ")
    Set rs = Nothing
    str = "select * from emp where e_no=" & key
    rs.Open str, adoconn, adOpenForwardOnly, adLockReadOnly
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
    Set rs = Nothing
    str = "select * from emp"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
End Sub

Private Sub cmdUpdate_Click()
    Dim ans As String
    ans = MsgBox("Do you really want to modify the current record?", vbExclamation + vbYesNo, "DELETE")
    If ans = vbYes Then
        rs.Update
        
        cmdFirst.Enabled = False
        cmdLast.Enabled = False
        cmdNext.Enabled = False
        cmdPrevious.Enabled = False
        cmdDelete.Enabled = False
        cmdSearch.Enabled = False
        cmdUpdate.Enabled = False
        cmdSave.Visible = True
        cmdAdd.Visible = False
    End If
End Sub

Private Sub Command1_Click()
MsgBox ("Created By: Mr. Jake Rodriguez Pomperada,MAED-IT April 6, 2009")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim str As String
    
    Set adoconn = Nothing
    adoconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=employee.mdb;Persist Security Info=False"
    str = "select * from emp"
    rs.Open str, adoconn, adOpenDynamic, adLockPessimistic
    rs.MoveFirst
    txtNo.Text = rs(0)
    txtName.Text = rs(1)
    txtCity.Text = rs(2)
    txtDob.Text = rs(4)
    txtPhone.Text = rs(3)
    txtdepartment.Text = rs(5)
End Sub

