VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Marksheet..."
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9525
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   9735
      Left            =   120
      TabIndex        =   9
      Top             =   -120
      Width           =   9255
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   8
         Top             =   9240
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5520
         TabIndex        =   50
         Top             =   9120
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         TabIndex        =   49
         Top             =   9120
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   7200
         TabIndex        =   48
         Top             =   8280
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7440
         TabIndex        =   4
         Top             =   7560
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7440
         TabIndex        =   3
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7440
         TabIndex        =   2
         Top             =   6600
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7440
         TabIndex        =   1
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   7440
         TabIndex        =   0
         Top             =   5640
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Percent"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   9120
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   9120
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Total Marks"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   8400
         Width           =   1575
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "500"
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
         Left            =   3360
         TabIndex        =   47
         Top             =   8400
         Width           =   1695
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "100"
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
         Left            =   3360
         TabIndex        =   46
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "100"
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
         Left            =   3360
         TabIndex        =   45
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Caption         =   "100"
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
         Left            =   3360
         TabIndex        =   44
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "100"
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
         Left            =   3360
         TabIndex        =   43
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "100"
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
         Left            =   3360
         TabIndex        =   42
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Caption         =   "33"
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
         Left            =   5280
         TabIndex        =   41
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "33"
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
         Left            =   5280
         TabIndex        =   40
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "33"
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
         Left            =   5280
         TabIndex        =   39
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Caption         =   "33"
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
         Left            =   5280
         TabIndex        =   38
         Top             =   7200
         Width           =   1695
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "33"
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
         Left            =   5280
         TabIndex        =   37
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   9000
         Y1              =   8160
         Y2              =   8160
      End
      Begin VB.Label Label16 
         Caption         =   "MATHEMATICS"
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
         Left            =   720
         TabIndex        =   36
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "CHEMISTRY"
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
         Left            =   720
         TabIndex        =   35
         Top             =   7200
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "PHYSICS"
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
         Left            =   720
         TabIndex        =   34
         Top             =   6720
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "ENGLISH"
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
         Left            =   720
         TabIndex        =   33
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "HINDI"
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
         Left            =   720
         TabIndex        =   32
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "OBTAIN"
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
         Left            =   7200
         TabIndex        =   31
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "MIN"
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
         Left            =   5280
         TabIndex        =   30
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "MAX"
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
         Left            =   3360
         TabIndex        =   29
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "MARKS"
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
         Left            =   5280
         TabIndex        =   28
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "SUBJECTS"
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
         Left            =   720
         TabIndex        =   27
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Line Line2 
         X1              =   5160
         X2              =   5160
         Y1              =   4920
         Y2              =   9000
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   9000
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Shape Shape5 
         Height          =   1215
         Left            =   240
         Top             =   4320
         Width           =   8775
      End
      Begin VB.Shape Shape4 
         Height          =   4095
         Left            =   3240
         Top             =   4920
         Width           =   3855
      End
      Begin VB.Shape Shape3 
         Height          =   4695
         Left            =   240
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Shape Shape2 
         Height          =   4695
         Left            =   240
         Top             =   4320
         Width           =   8775
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "REGISTRATION NO."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   26
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Caption         =   "STATUS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "131015245"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "REGULAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   23
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Line Line13 
         X1              =   5520
         X2              =   5520
         Y1              =   2040
         Y2              =   3120
      End
      Begin VB.Line Line12 
         X1              =   2760
         X2              =   2760
         Y1              =   2040
         Y2              =   3120
      End
      Begin VB.Line Line11 
         X1              =   960
         X2              =   8280
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Shape Shape1 
         Height          =   1095
         Left            =   960
         Top             =   2040
         Width           =   7335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "MARKSHEET"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "AARYAN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   21
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Caption         =   "CHHATTISGARH BOARD OF SECONDARY EDUCATION,RAIPUR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   9015
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "HIGHER SECONDRY SCHOOL CERTIFICATE EXAMINATION (10+2)2014"
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
         Left            =   480
         TabIndex        =   17
         Top             =   720
         Width           =   8415
      End
      Begin VB.Label Label24 
         Caption         =   "S.No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6120
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "FATHER'S NAME SRI"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "ROLL NUMBER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "MARCH, 2014"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "A08/131271/007"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "10132757"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "KUMAR"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   3720
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text6.Text = Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
End Sub

Private Sub Command2_Click()
Text7.Text = Val(Text6.Text) / 500 * 100
End Sub

Private Sub Command3_Click()
Dim n As Integer
n = Val(Text7.Text)
If n >= 90 Then
    Text8.Text = "A+"
ElseIf n >= 75 And n < 90 Then
    Text8.Text = "A"
ElseIf n >= 60 And n < 75 Then
    Text8.Text = "B"
ElseIf n >= 45 And n < 60 Then
    Text8.Text = "C"
Else
    Text8.Text = "F"
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
