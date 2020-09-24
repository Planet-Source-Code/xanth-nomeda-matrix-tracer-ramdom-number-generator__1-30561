VERSION 5.00
Begin VB.Form frmOrigonal 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Matrix Tracer... by Nicholas Romanelli"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer9 
      Interval        =   500
      Left            =   4320
      Top             =   240
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset Trace"
      Height          =   315
      Left            =   2760
      TabIndex        =   117
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9120
      TabIndex        =   116
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   120
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9720
      Top             =   1200
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8280
      Top             =   1200
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   1200
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1200
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8880
      TabIndex        =   4
      Text            =   "0"
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   315
      Left            =   8160
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   315
      Left            =   5520
      TabIndex        =   1
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Trace"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "!!Trace Complete!!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   5040
      TabIndex        =   126
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Line Line7 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   9720
      X2              =   9720
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Line Line6 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   8280
      X2              =   8280
      Y1              =   1440
      Y2              =   5520
   End
   Begin VB.Line Line5 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   6840
      X2              =   6840
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Line Line4 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Line Line3 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   3720
      X2              =   3720
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1800
      TabIndex        =   125
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1560
      TabIndex        =   124
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1320
      TabIndex        =   123
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1080
      TabIndex        =   122
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   840
      TabIndex        =   121
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   600
      TabIndex        =   120
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   360
      TabIndex        =   119
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   118
      Top             =   360
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderWidth     =   11
      DrawMode        =   1  'Blackness
      Visible         =   0   'False
      X1              =   1920
      X2              =   1920
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderWidth     =   11
      Visible         =   0   'False
      X1              =   650
      X2              =   650
      Y1              =   1560
      Y2              =   5640
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   9360
      TabIndex        =   115
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   9360
      TabIndex        =   114
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   9360
      TabIndex        =   113
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   9360
      TabIndex        =   112
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   9360
      TabIndex        =   111
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   9360
      TabIndex        =   110
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   9360
      TabIndex        =   109
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   9360
      TabIndex        =   108
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   107
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   9360
      TabIndex        =   106
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   105
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   104
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   9360
      TabIndex        =   103
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   102
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   101
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   7920
      TabIndex        =   100
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   99
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   7920
      TabIndex        =   98
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   7920
      TabIndex        =   97
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   7920
      TabIndex        =   96
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   95
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   7920
      TabIndex        =   94
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   93
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   92
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   91
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   90
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   89
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   88
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   87
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   86
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   6360
      TabIndex        =   85
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   6360
      TabIndex        =   84
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   6360
      TabIndex        =   83
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   6360
      TabIndex        =   82
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   6360
      TabIndex        =   81
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   6360
      TabIndex        =   80
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   6360
      TabIndex        =   79
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   6360
      TabIndex        =   78
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   6360
      TabIndex        =   77
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   6360
      TabIndex        =   76
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   6360
      TabIndex        =   75
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   6360
      TabIndex        =   74
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   6360
      TabIndex        =   73
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   72
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   71
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   4680
      TabIndex        =   70
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   4680
      TabIndex        =   69
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   4680
      TabIndex        =   68
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   67
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   66
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   65
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   64
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   63
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   62
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   61
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   60
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   59
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   58
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   57
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   56
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   3120
      TabIndex        =   55
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   3120
      TabIndex        =   54
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   53
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   3120
      TabIndex        =   52
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   3120
      TabIndex        =   51
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   3120
      TabIndex        =   50
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   3120
      TabIndex        =   49
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   3120
      TabIndex        =   48
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   47
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   46
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   45
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   44
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   42
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   41
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   1560
      TabIndex        =   40
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   39
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   1560
      TabIndex        =   38
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   1560
      TabIndex        =   37
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   1560
      TabIndex        =   36
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   35
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   1560
      TabIndex        =   34
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   1560
      TabIndex        =   33
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   1560
      TabIndex        =   32
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   1560
      TabIndex        =   31
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   1560
      TabIndex        =   30
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   29
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   28
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   27
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   26
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   9360
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   24
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   23
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   22
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   21
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   19
      Top             =   5280
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   18
      Top             =   5040
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0000000"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmOrigonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*************************************************
'*This is the origonal file by Nicholas Romanelli*
'*************************************************

Public Function rndstr(strlen As Integer) As String
    Dim charpos As Integer
    Dim cstrlen As Integer
    Dim rndstring As String
    Dim chars As String
    chars = "123456789"
    cstrlen = 0
    Randomize


    Do Until cstrlen = strlen
        charpos = Int((Len(chars) * Rnd) + 1)
        rndstring = rndstring & Mid(chars, charpos, 1)
        cstrlen = cstrlen + 1
    Loop
    rndstr = rndstring
End Function

Private Sub Command1_Click()
Text1.Text = "0"
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
Timer4.Enabled = True
Timer5.Enabled = True
Timer6.Enabled = True
Timer7.Enabled = True
Timer8.Enabled = True
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = False
Timer8.Enabled = False
Label8.Caption = "0"
Label9.Caption = "0"
Label10.Caption = "0"
Label12.Caption = "0"
Label13.Caption = "0"
Label14.Caption = "0"
Label15.Caption = "0"
Line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
For c = 0 To 15
Label1(c).Caption = "0000000"
Label2(c).Caption = "0000000"
Label3(c).Caption = "0000000"
Label4(c).Caption = "0000000"
Label5(c).Caption = "0000000"
Label6(c).Caption = "0000000"
Label7(c).Caption = "0000000"
Next c
End Sub

Private Sub Timer1_Timer()
For c = 0 To 15
Label1(c).Caption = rndstr(7)
Next c
End Sub

Private Sub Timer2_Timer()
For c = 0 To 15
Label2(c).Caption = rndstr(7)
Next c

End Sub

Private Sub Timer3_Timer()
For c = 0 To 15
Label3(c).Caption = rndstr(7)
Next c

End Sub

Private Sub Timer4_Timer()
For c = 0 To 15
Label4(c).Caption = rndstr(7)
Next c

End Sub

Private Sub Timer5_Timer()
For c = 0 To 15
Label5(c).Caption = rndstr(7)
Next c

End Sub

Private Sub Timer6_Timer()
For c = 0 To 15
Label6(c).Caption = rndstr(7)
Next c


End Sub

Private Sub Timer7_Timer()
For c = 0 To 15
Label7(c).Caption = rndstr(7)
Next c

End Sub

Private Sub Timer8_Timer()
Text2.Text = rndstr(1)
Text1.Text = Val(Text1.Text) + 1
If Text1.Text = "3" Then
  If Text2.Text = "1" Then Line1.X1 = "200"
  If Text2.Text = "1" Then Line1.X2 = "200"
  If Text2.Text = "2" Then Line1.X1 = "360"
  If Text2.Text = "2" Then Line1.X2 = "360"
  If Text2.Text = "3" Then Line1.X1 = "480"
  If Text2.Text = "3" Then Line1.X2 = "480"
  If Text2.Text = "4" Then Line1.X1 = "675"
  If Text2.Text = "4" Then Line1.X2 = "675"
  If Text2.Text = "5" Then Line1.X1 = "800"
  If Text2.Text = "5" Then Line1.X2 = "800"
  If Text2.Text = "6" Then Line1.X1 = "960"
  If Text2.Text = "6" Then Line1.X2 = "960"
  If Text2.Text = "7" Then Line1.X1 = "1080"
  If Text2.Text = "7" Then Line1.X2 = "1080"
  Line1.Visible = True
  Label8.Caption = rndstr(1)
End If
If Text1.Text = "6" Then
  If Text2.Text = "1" Then Line2.X1 = "1600"
  If Text2.Text = "1" Then Line2.X2 = "1600"
  If Text2.Text = "2" Then Line2.X1 = "1800"
  If Text2.Text = "2" Then Line2.X2 = "1800"
  If Text2.Text = "3" Then Line2.X1 = "1920"
  If Text2.Text = "3" Then Line2.X2 = "1920"
  If Text2.Text = "4" Then Line2.X1 = "2100"
  If Text2.Text = "4" Then Line2.X2 = "2100"
  If Text2.Text = "5" Then Line2.X1 = "2250"
  If Text2.Text = "5" Then Line2.X2 = "2250"
  If Text2.Text = "6" Then Line2.X1 = "2400"
  If Text2.Text = "6" Then Line2.X2 = "2400"
  If Text2.Text = "7" Then Line2.X1 = "2520"
  If Text2.Text = "7" Then Line2.X2 = "2520"
  Line2.Visible = True
  Label9.Caption = rndstr(1)
End If
If Text1.Text = "9" Then
  If Text2.Text = "1" Then Line3.X1 = "3200"
  If Text2.Text = "1" Then Line3.X2 = "3200"
  If Text2.Text = "2" Then Line3.X1 = "3360"
  If Text2.Text = "2" Then Line3.X2 = "3360"
  If Text2.Text = "3" Then Line3.X1 = "3480"
  If Text2.Text = "3" Then Line3.X2 = "3480"
  If Text2.Text = "4" Then Line3.X1 = "3650"
  If Text2.Text = "4" Then Line3.X2 = "3650"
  If Text2.Text = "5" Then Line3.X1 = "3800"
  If Text2.Text = "5" Then Line3.X2 = "3800"
  If Text2.Text = "6" Then Line3.X1 = "3960"
  If Text2.Text = "6" Then Line3.X2 = "3960"
  If Text2.Text = "7" Then Line3.X1 = "4080"
  If Text2.Text = "7" Then Line3.X2 = "4080"
  Line3.Visible = True
  Label10.Caption = rndstr(1)
End If
If Text1.Text = "12" Then
  If Text2.Text = "1" Then Line4.X1 = "4750"
  If Text2.Text = "1" Then Line4.X2 = "4750"
  If Text2.Text = "2" Then Line4.X1 = "4920"
  If Text2.Text = "2" Then Line4.X2 = "4920"
  If Text2.Text = "3" Then Line4.X1 = "5040"
  If Text2.Text = "3" Then Line4.X2 = "5040"
  If Text2.Text = "4" Then Line4.X1 = "5200"
  If Text2.Text = "4" Then Line4.X2 = "5200"
  If Text2.Text = "5" Then Line4.X1 = "5350"
  If Text2.Text = "5" Then Line4.X2 = "5350"
  If Text2.Text = "6" Then Line4.X1 = "5520"
  If Text2.Text = "6" Then Line4.X2 = "5520"
  If Text2.Text = "7" Then Line4.X1 = "5640"
  If Text2.Text = "7" Then Line4.X2 = "5640"
  Line4.Visible = True
  Label12.Caption = rndstr(1)
End If
If Text1.Text = "15" Then
  If Text2.Text = "1" Then Line5.X1 = "6400"
  If Text2.Text = "1" Then Line5.X2 = "6400"
  If Text2.Text = "2" Then Line5.X1 = "6600"
  If Text2.Text = "2" Then Line5.X2 = "6600"
  If Text2.Text = "3" Then Line5.X1 = "6720"
  If Text2.Text = "3" Then Line5.X2 = "6720"
  If Text2.Text = "4" Then Line5.X1 = "6900"
  If Text2.Text = "4" Then Line5.X2 = "6900"
  If Text2.Text = "5" Then Line5.X1 = "7000"
  If Text2.Text = "5" Then Line5.X2 = "7000"
  If Text2.Text = "6" Then Line5.X1 = "7200"
  If Text2.Text = "6" Then Line5.X2 = "7200"
  If Text2.Text = "7" Then Line5.X1 = "7320"
  If Text2.Text = "7" Then Line5.X2 = "7320"
  Line5.Visible = True
  Label13.Caption = rndstr(1)
End If
If Text1.Text = "18" Then
  If Text2.Text = "1" Then Line6.X1 = "8000"
  If Text2.Text = "1" Then Line6.X2 = "8000"
  If Text2.Text = "2" Then Line6.X1 = "8160"
  If Text2.Text = "2" Then Line6.X2 = "8160"
  If Text2.Text = "3" Then Line6.X1 = "8280"
  If Text2.Text = "3" Then Line6.X2 = "8280"
  If Text2.Text = "4" Then Line6.X1 = "8450"
  If Text2.Text = "4" Then Line6.X2 = "8450"
  If Text2.Text = "5" Then Line6.X1 = "8600"
  If Text2.Text = "5" Then Line6.X2 = "8600"
  If Text2.Text = "6" Then Line6.X1 = "8760"
  If Text2.Text = "6" Then Line6.X2 = "8760"
  If Text2.Text = "7" Then Line6.X1 = "8880"
  If Text2.Text = "7" Then Line6.X2 = "8880"
  Line6.Visible = True
  Label14.Caption = rndstr(1)
End If
If Text1.Text = "21" Then
  If Text2.Text = "1" Then Line7.X1 = "9425"
  If Text2.Text = "1" Then Line7.X2 = "9425"
  If Text2.Text = "2" Then Line7.X1 = "9600"
  If Text2.Text = "2" Then Line7.X2 = "9600"
  If Text2.Text = "3" Then Line7.X1 = "9720"
  If Text2.Text = "3" Then Line7.X2 = "9720"
  If Text2.Text = "4" Then Line7.X1 = "9875"
  If Text2.Text = "4" Then Line7.X2 = "9875"
  If Text2.Text = "5" Then Line7.X1 = "10000"
  If Text2.Text = "5" Then Line7.X2 = "10000"
  If Text2.Text = "6" Then Line7.X1 = "10200"
  If Text2.Text = "6" Then Line7.X2 = "10200"
  If Text2.Text = "7" Then Line7.X1 = "10320"
  If Text2.Text = "7" Then Line7.X2 = "10320"
  Line7.Visible = True
  Label15.Caption = rndstr(1)
End If
End Sub

Private Sub Timer9_Timer()
If Label15.Caption <> "0" Then Label16.Visible = True
If Label15.Caption = "0" Then Label16.Visible = False
End Sub
