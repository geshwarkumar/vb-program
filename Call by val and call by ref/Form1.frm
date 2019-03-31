VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Swap By Refrence"
      Height          =   1095
      Left            =   1440
      TabIndex        =   29
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Swap By Value"
      Height          =   1095
      Left            =   1440
      TabIndex        =   28
      Top             =   4440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      Caption         =   "Swap By Refrence"
      Height          =   3495
      Left            =   4320
      TabIndex        =   2
      Top             =   3960
      Width           =   5055
      Begin VB.Frame Frame7 
         Caption         =   "Swapped Value"
         Height          =   2295
         Left            =   2880
         TabIndex        =   8
         Top             =   840
         Width           =   1815
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   480
            TabIndex        =   26
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   480
            TabIndex        =   25
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            TabIndex        =   24
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "A"
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
            TabIndex        =   23
            Top             =   600
            Width           =   135
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Passed Value"
         Height          =   2295
         Left            =   600
         TabIndex        =   7
         Top             =   840
         Width           =   1815
         Begin VB.TextBox Text8 
            Height          =   495
            Left            =   480
            TabIndex        =   22
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   480
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "A"
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
            Top             =   600
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Swap By Value"
      Height          =   3495
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   5055
      Begin VB.Frame Frame5 
         Caption         =   "Swapped Value"
         Height          =   2295
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   360
            TabIndex        =   18
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   495
            Left            =   360
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "A"
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
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   135
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Passed Value"
         Height          =   2295
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         Begin VB.TextBox Text4 
            Height          =   495
            Left            =   480
            TabIndex        =   14
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   480
            TabIndex        =   13
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "B"
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
            TabIndex        =   12
            Top             =   1440
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "A"
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
            TabIndex        =   11
            Top             =   600
            Width           =   135
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   3495
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Pass Value"
         Height          =   615
         Left            =   480
         TabIndex        =   27
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "B"
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
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "A"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Integer
Dim y As Integer
Dim temp As Integer
Private Sub Command3_Click()
Call swap_rf(x, y)
End Sub
Private Sub Command2_Click()
Call swap_vl(x, y)
End Sub
Private Sub Command1_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
End Sub
Function swap_vl(ByVal a As Integer, b As Integer)
Text3.Text = x
Text4.Text = y
temp = a
a = b
b = temp
Text5.Text = a
Text6.Text = b
End Function
Function swap_rf(ByRef a As Integer, b As Integer)
Text7.Text = x
Text8.Text = y
temp = a
a = b
b = temp
Text9 = a
Text10 = b
End Function
