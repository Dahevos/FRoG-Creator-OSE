VERSION 5.00
Begin VB.Form frmbank 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Banque"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmbank.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   9360
      TabIndex        =   174
      Top             =   3290
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   9360
      TabIndex        =   173
      Top             =   2930
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   9360
      TabIndex        =   172
      Top             =   2570
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   9360
      TabIndex        =   171
      Top             =   2210
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   9360
      TabIndex        =   170
      Top             =   1850
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   9360
      TabIndex        =   169
      Top             =   1490
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   9360
      TabIndex        =   168
      Top             =   1140
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   9360
      TabIndex        =   167
      Top             =   770
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   9360
      TabIndex        =   166
      Top             =   4370
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   9360
      TabIndex        =   165
      Top             =   4010
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   9360
      TabIndex        =   164
      Top             =   3650
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   9360
      TabIndex        =   163
      Top             =   3290
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   9360
      TabIndex        =   162
      Top             =   2930
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   9360
      TabIndex        =   161
      Top             =   2570
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   9360
      TabIndex        =   160
      Top             =   2210
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   9360
      TabIndex        =   159
      Top             =   1850
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   9360
      TabIndex        =   158
      Top             =   1490
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   9360
      TabIndex        =   157
      Top             =   1140
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   9360
      TabIndex        =   156
      Top             =   770
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   9240
      TabIndex        =   155
      Top             =   5400
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   7470
      TabIndex        =   154
      Top             =   3290
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   7470
      TabIndex        =   153
      Top             =   2930
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   7470
      TabIndex        =   152
      Top             =   2570
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   7470
      TabIndex        =   151
      Top             =   2210
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   7470
      TabIndex        =   150
      Top             =   1850
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   7470
      TabIndex        =   149
      Top             =   1490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Arc"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   7470
      TabIndex        =   148
      Top             =   1140
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Epée"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   7470
      TabIndex        =   147
      Top             =   770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   146
      Top             =   5400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   7470
      TabIndex        =   145
      Top             =   4370
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   7470
      TabIndex        =   144
      Top             =   4010
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   7470
      TabIndex        =   143
      Top             =   3650
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   7470
      TabIndex        =   142
      Top             =   3290
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   7470
      TabIndex        =   141
      Top             =   2930
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   7470
      TabIndex        =   140
      Top             =   2570
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   7470
      TabIndex        =   139
      Top             =   2210
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   7470
      TabIndex        =   138
      Top             =   1850
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   7470
      TabIndex        =   137
      Top             =   1490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Arc"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   7470
      TabIndex        =   136
      Top             =   1140
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Epée"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   7470
      TabIndex        =   135
      Top             =   770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   3000
      TabIndex        =   134
      Top             =   1835
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   3000
      TabIndex        =   133
      Top             =   1490
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   3000
      TabIndex        =   132
      Top             =   1140
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   3000
      TabIndex        =   131
      Top             =   765
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   24
      Left            =   1140
      TabIndex        =   130
      Top             =   1835
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   1140
      TabIndex        =   129
      Top             =   1490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Boule de cristal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   1140
      TabIndex        =   128
      Top             =   1140
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   1140
      TabIndex        =   127
      Top             =   770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Haut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2400
      TabIndex        =   126
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   125
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2850
      TabIndex        =   124
      Top             =   4440
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   121
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   9480
      TabIndex        =   119
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   9480
      TabIndex        =   118
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   9480
      TabIndex        =   117
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   9480
      TabIndex        =   116
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   9480
      TabIndex        =   115
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   9480
      TabIndex        =   114
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   9480
      TabIndex        =   113
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   9480
      TabIndex        =   112
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   9480
      TabIndex        =   111
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   19
      Left            =   9480
      TabIndex        =   110
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   18
      Left            =   9480
      TabIndex        =   109
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   17
      Left            =   9480
      TabIndex        =   108
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   16
      Left            =   9480
      TabIndex        =   107
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   15
      Left            =   9480
      TabIndex        =   106
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   9480
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   9480
      TabIndex        =   104
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   9480
      TabIndex        =   103
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 10 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   6720
      TabIndex        =   102
      Top             =   4010
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 9 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6720
      TabIndex        =   101
      Top             =   3650
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 8 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   6720
      TabIndex        =   100
      Top             =   3290
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 7 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   6720
      TabIndex        =   99
      Top             =   2930
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 6 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6720
      TabIndex        =   98
      Top             =   2570
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 5 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6720
      TabIndex        =   97
      Top             =   2210
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 4 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6720
      TabIndex        =   96
      Top             =   1850
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 3 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   95
      Top             =   1490
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 1 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6720
      TabIndex        =   94
      Top             =   770
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 2 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   93
      Top             =   1130
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 11 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   6720
      TabIndex        =   92
      Top             =   4370
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Epée"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   7470
      TabIndex        =   91
      Top             =   770
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   7460
      Top             =   750
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Arc"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   7470
      TabIndex        =   90
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   7470
      TabIndex        =   89
      Top             =   1490
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   7470
      TabIndex        =   88
      Top             =   1850
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   7470
      TabIndex        =   87
      Top             =   2210
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   7470
      TabIndex        =   86
      Top             =   2570
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7470
      TabIndex        =   85
      Top             =   2930
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   7470
      TabIndex        =   84
      Top             =   3290
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   7470
      TabIndex        =   83
      Top             =   3650
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   7470
      TabIndex        =   82
      Top             =   4010
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   7470
      TabIndex        =   81
      Top             =   4370
      Width           =   1815
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   9360
      TabIndex        =   80
      Top             =   4370
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   9360
      TabIndex        =   79
      Top             =   770
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   9360
      TabIndex        =   78
      Top             =   1130
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   9360
      TabIndex        =   77
      Top             =   1490
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   76
      Top             =   1850
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   75
      Top             =   2210
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   9360
      TabIndex        =   74
      Top             =   2570
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   73
      Top             =   2930
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   9360
      TabIndex        =   72
      Top             =   3290
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   9360
      TabIndex        =   71
      Top             =   3650
      Width           =   1425
   End
   Begin VB.Label valcof 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   9360
      TabIndex        =   70
      Top             =   4010
      Width           =   1425
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   22
      Left            =   9480
      TabIndex        =   69
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   21
      Left            =   9480
      TabIndex        =   68
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   20
      Left            =   9480
      TabIndex        =   67
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   23
      Left            =   9480
      TabIndex        =   66
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   24
      Left            =   9480
      TabIndex        =   65
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   25
      Left            =   9480
      TabIndex        =   64
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   26
      Left            =   9480
      TabIndex        =   63
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   27
      Left            =   9480
      TabIndex        =   62
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   28
      Left            =   9480
      TabIndex        =   61
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   29
      Left            =   9480
      TabIndex        =   60
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label numcof 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   30
      Left            =   9480
      TabIndex        =   59
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   58
      Top             =   770
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   57
      Top             =   1140
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   56
      Top             =   1490
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   55
      Top             =   1835
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   54
      Top             =   2205
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   53
      Top             =   2565
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   52
      Top             =   2925
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   51
      Top             =   3285
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   50
      Top             =   3645
      Width           =   1545
   End
   Begin VB.Label valinv 
      BackStyle       =   0  'Transparent
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   49
      Top             =   4005
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1140
      TabIndex        =   48
      Top             =   770
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1140
      TabIndex        =   47
      Top             =   2925
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Boule de cristal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1140
      TabIndex        =   46
      Top             =   1140
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1140
      TabIndex        =   45
      Top             =   1835
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   1140
      TabIndex        =   44
      Top             =   2565
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   1140
      TabIndex        =   43
      Top             =   2205
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   1140
      TabIndex        =   42
      Top             =   3645
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1140
      TabIndex        =   41
      Top             =   3285
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   1140
      TabIndex        =   40
      Top             =   4005
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1140
      TabIndex        =   39
      Top             =   1490
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 2 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   38
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 3 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   36
      Top             =   1490
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 4 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   35
      Top             =   1835
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 5 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   34
      Top             =   2205
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 6 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   33
      Top             =   2565
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 7 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   32
      Top             =   2925
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 8 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   31
      Top             =   3285
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 9 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   30
      Top             =   3645
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 10 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   29
      Top             =   4005
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1120
      Top             =   750
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   1140
      TabIndex        =   28
      Top             =   4005
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   1140
      TabIndex        =   27
      Top             =   770
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   1140
      TabIndex        =   26
      Top             =   1140
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   1140
      TabIndex        =   25
      Top             =   1490
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   1140
      TabIndex        =   24
      Top             =   1835
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   1140
      TabIndex        =   23
      Top             =   2205
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   1140
      TabIndex        =   22
      Top             =   2565
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Boule de cristal"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   1140
      TabIndex        =   21
      Top             =   2925
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   1140
      TabIndex        =   20
      Top             =   3285
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   1140
      TabIndex        =   19
      Top             =   3645
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   3000
      TabIndex        =   18
      Top             =   4005
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3000
      TabIndex        =   17
      Top             =   3645
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   3000
      TabIndex        =   16
      Top             =   3285
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3000
      TabIndex        =   15
      Top             =   2925
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   3000
      TabIndex        =   14
      Top             =   2565
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   13
      Top             =   2205
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   3000
      TabIndex        =   12
      Top             =   1835
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   11
      Top             =   1490
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   10
      Top             =   1140
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label valinv 
      Caption         =   "Valeur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   9
      Top             =   770
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label metreinv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mettre dans l'inventaire"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4800
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label metrecofre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mettre dans le coffre"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   4800
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label jeter 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jeter"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9240
      TabIndex        =   6
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Haut2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Bas2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label annuler 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Annuler"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label OK 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OK"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Bas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label jinv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Jeter"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slot 1 :"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   37
      Top             =   770
      Width           =   615
   End
   Begin VB.Label inve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventaire :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   122
      Top             =   500
      Width           =   795
   End
   Begin VB.Label coffre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coffre :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6720
      TabIndex        =   123
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "frmbank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Pic
Dim Num2
Private Sub annuler_Click()
Pic = vbNullString
Num2 = vbNullString
frmMirage.txtQ.Visible = False
Unload Me
End Sub

Private Sub Bas_Click()
If label1(1).Visible = True Then
For ml = 1 To 10
label1(ml).Visible = False
valinv(ml).Visible = False
Label2(ml).Visible = True
Next ml

For po = 11 To 20
label1(po).Visible = True
valinv(po).Visible = True
Next po

lop = 11
For lpo = 1 To 10
Label2(lpo).Caption = "slot " & lop & " :"
lop = lop + 1
Next lpo

Else


For lm = 11 To 20
label1(lm).Visible = False
valinv(lm).Visible = False
Next lm

For pot = 21 To 24
label1(pot).Visible = True
valinv(pot).Visible = True
Next pot

lops = 21
For lpos = 1 To 4
Label2(lpos).Caption = "slot " & lops & " :"
lops = lops + 1
Next lpos

For opss = 5 To 10
Label2(opss).Visible = False
Next opss

End If

'Call Form_Load

End Sub

Private Sub Bas2_Click()

If Label3(1).Visible = True Then

For ml = 1 To 11
Label3(ml).Visible = False
valcof(ml).Visible = False
Label4(ml).Visible = True
Next ml

For po = 12 To 22
Label3(po).Visible = True
valcof(po).Visible = True
Next po

lop = 12
For lpo = 1 To 11
Label4(lpo).Caption = "slot " & lop & " :"
lop = lop + 1
Next lpo

Else

For lm = 12 To 22
Label3(lm).Visible = False
valcof(lm).Visible = False
Next lm

For pot = 23 To 30
Label3(pot).Visible = True
valcof(pot).Visible = True
Next pot

lops = 23
For lpos = 1 To 8
Label4(lpos).Caption = "slot " & lops & " :"
lops = lops + 1
Next lpos

For opss = 9 To 11
Label4(opss).Visible = False
Next opss
End If

End Sub

Private Sub Haut_Click()
If label1(11).Visible = True Then
For W = 1 To 10
label1(W).Visible = True
valinv(W).Visible = True
Label2(W).Visible = True
Next W

For ty = 11 To 20
label1(ty).Visible = False
valinv(ty).Visible = False
Next ty

For lp = 1 To 10
Label2(lp).Caption = "slot " & lp & " :"
Next lp

End If

If label1(21).Visible = True Then
For Ws = 11 To 20
label1(Ws).Visible = True
valinv(Ws).Visible = True
Next Ws

For tys = 21 To 24
label1(tys).Visible = False
valinv(tys).Visible = False
Next tys

lpso = 11
For lps = 1 To 10
Label2(lps).Caption = "slot " & lpso & " :"
Label2(lps).Visible = True
lpso = lpso + 1
Next lps

End If
'Call Form_Load

End Sub

Private Sub Haut2_Click()
If Label3(12).Visible = True Then
For W = 1 To 11
Label3(W).Visible = True
valcof(W).Visible = True
Label4(W).Visible = True
Next W

For ty = 12 To 22
Label3(ty).Visible = False
valcof(ty).Visible = False
Next ty

For lp = 1 To 11
Label4(lp).Caption = "slot " & lp & " :"
Next lp

End If

If Label3(23).Visible = True Then
For Ws = 12 To 22
Label3(Ws).Visible = True
valcof(Ws).Visible = True
Next Ws

For tys = 23 To 30
Label3(tys).Visible = False
valcof(tys).Visible = False
Next tys

lpso = 12
For lps = 1 To 11
Label4(lps).Caption = "slot " & lpso & " :"
Label4(lps).Visible = True
lpso = lpso + 1
Next lps

End If
End Sub

Public Sub jeter_Click()
Dim Packet As String
If Num2 = 0 Then variable = MsgBox("Veuillez sélectioner un slot dans le coffre!!", vbCritical, "Erreur")
If Num2 = 0 Then GoTo jt:
If Num2 = vbNullString Then variable = MsgBox("Veuillez sélectioner un slot dans le coffre!!", vbCritical, "Erreur")
If Num2 = vbNullString Then GoTo jt:
If Label3(Num2).Caption = vbNullString Then variable = MsgBox("Aucun objet dans le slot" & Num2 & " du coffre!!!", vbCritical, "Erreur")
If Label3(Num2).Caption = vbNullString Then GoTo jt:
cont = MsgBox("Voulez vous vraiment jeter " & Mid(valcof(Num2).Caption, 9) & Trim(Label3(Num2).Caption) & " du coffre?? il sera supprimé définitivement!!", vbYesNo, "Demande")
If cont = vbYes Then

Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)
Packet = vbNullString
End If


Call Form_Load
jt:
End Sub

Private Sub jinv_Click()
Dim Packet As String
If Pic = 0 Then variable = MsgBox("Veuillez sélectioner un slot dans l'inventaire!!", vbCritical, "Erreur")
If Pic = 0 Then GoTo jti:
If Pic = vbNullString Then variable = MsgBox("Veuillez sélectioner un slot dans l'inventaire!!", vbCritical, "Erreur")
If Pic = vbNullString Then GoTo jti:
If label1(Pic).Caption = vbNullString Then variable = MsgBox("Aucun objet dans le slot" & Pic & " de l'inventaire!!!", vbCritical, "Erreur")
If label1(Pic).Caption = vbNullString Then GoTo jti:
cont = MsgBox("Voulez vous vraiment jeter " & Mid(valinv(Pic).Caption, 9) & Trim(label1(Pic).Caption) & " de l'inventaire?? il sera supprimer définitivement!!", vbYesNo, "Demande")

If cont = vbYes Then

Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)

End If

Call SendData("REFRESH" & SEP_CHAR & END_CHAR)

Call Form_Load

jti:
Call Form_Load
End Sub

Private Sub Label1_Click(index As Integer)
Pic = index
num = index
Shape3.Visible = True
Shape3.Left = label1(num).Left - 15
Shape3.Top = label1(num).Top - 20
End Sub

Private Sub Label3_Click(index As Integer)
Num2 = index
Shape2.Visible = True
Shape2.Left = Label3(Num2).Left - 15
Shape2.Top = Label3(Num2).Top - 15
End Sub

Private Sub ok_Click()
Pic = vbNullString
Num2 = vbNullString
frmMirage.txtQ.Visible = False
Unload Me
End Sub
Public Sub Form_Load()
Dim Packet
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExiste("GUI\Bank" & Ending) Then frmbank.Picture = LoadPicture(App.Path & "\GUI\Bank" & Ending)
    Next i
    
inve.Caption = "Inventaire de " & GetPlayerName(MyIndex) & " :"
coffre.Caption = "Coffre de " & GetPlayerName(MyIndex) & " :"
W:
For m = 1 To 24
label1(m).Caption = vbNullString
valinv(m).Caption = "Nombre :"
Next m

For d = 1 To 30
Label3(d).Caption = vbNullString
valcof(d).Caption = "Nombre :"
Next d

For i = 1 To 24
Inum = GetPlayerInvItemNum(MyIndex, i)
If Inum = 0 Then GoTo i:
If Inum = "0" Or Inum = vbNullString Then GoTo i:
label1(i).Caption = Item(Inum).name
i:
Next i

For a = 1 To 24
Va = GetPlayerInvItemValue(MyIndex, a)
Inum2 = GetPlayerInvItemNum(MyIndex, a)
If Va = 0 And Inum2 = 0 Then GoTo a:
If Va = 0 And Inum2 > 0 Then Va = 1
If Va < 0 Then Va = vbNullString
If Va > 9999999999# Then Va = "+de9Milliard"
valinv(a).Caption = "Nombre :" & Va
a:
Next a

For loui = 1 To 24
valinv(loui).BackStyle = 0
valinv(loui).Appearance = 0
valinv(loui).BorderStyle = 1
Next loui

For louis = 1 To 30
valcof(louis).BackStyle = 0
valcof(louis).Appearance = 0
valcof(louis).BorderStyle = 1
Next louis

Packet = "COFFREITEM" & SEP_CHAR & END_CHAR

Call SendData(Packet)

End Sub
Private Sub metreinv_Click()
Dim Packet As String
Dim ItemNume As Long
Dim ItemVale As Long
Dim ItemDura As Long

If Pic = vbNullString And Num2 = vbNullString Then variable = MsgBox("Aucun slot sélectioné!!!", vbCritical, "Erreur")
If Pic = vbNullString And Num2 = vbNullString Then GoTo apr:
If Pic = vbNullString Then variable = MsgBox("Aucun slot sélectioné dans l'inventaire!!!", vbCritical, "Erreur")
If Pic = vbNullString Then GoTo apr:
If Num2 = vbNullString Then variable = MsgBox("Veuillez séléctioner un slot dans le coffre S.V.P!!", vbCritical, "Erreur")
If Num2 = vbNullString Then GoTo apr:
v2 = Mid(valcof(Num2).Caption, 9, 100)
Va10 = Val(v2)
Inum10 = Val(Trim(numcof(Num2).Caption)) 'GetPlayerInvItemNum(MyIndex, Pic)
If v2 = vbNullString And Label3(Num2).Caption = vbNullString Then variable = MsgBox("Aucun objet dans le slot" & Num2 & " du coffre!!", vbCritical, "Erreur")
If v2 = vbNullString And Label3(Num2).Caption = vbNullString Then GoTo apr:
If label1(Pic).Caption = vbNullString Then
Else
If Item(Inum10).Type = 12 And Label3(Num2).Caption = label1(Pic).Caption Then
Else
variable = MsgBox("Il y a déjà un objet dans le slot" & Pic & " de l'inventaire!!", vbCritical, "Erreur")
GoTo apr:
End If
End If
If Num2 <= 0 Then variable = MsgBox("Veuillez séléctioner un slot dans le coffre S.V.P!!", vbCritical, "Erreur")
If Num2 <= 0 Then GoTo apr:
ini3 = InputBox("Conbiens d'objet(s) voulez vous métre dans l'inventaire?", "Demande")
ini4 = Val(ini3)
If ini4 > Va10 Then variable = MsgBox("Valeur supérieur au nombre d'objet!!", vbCritical, "Erreur")
If ini4 > Va10 Then GoTo apr:
If ini3 = vbNullString Then GoTo apr:
If ini3 = "0" Then GoTo apr:
If ini4 = 0 Then variable = MsgBox("Veuillez saisir un nombre svp!!!", vbCritical, "Erreur")
If ini4 = 0 Then GoTo apr:

If Item(Inum10).Type = 12 And Label3(Num2).Caption = label1(Pic).Caption Or label1(Pic).Caption = vbNullString Then
ItemNume = Trim(numcof(Num2).Caption)
ItemVale = Val(Mid(valinv(Pic).Caption, 9, 100)) + ini4
ItemDura = Item(ItemNume).Data1

Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & ItemNume & SEP_CHAR & ItemVale & SEP_CHAR & ItemDura & SEP_CHAR & END_CHAR

Call SendData(Packet)

If Va10 - ini4 > 0 Then
Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & ItemNume & SEP_CHAR & Va10 - ini4 & SEP_CHAR & END_CHAR

Call SendData(Packet)

Else

Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)

End If

valcof(Num2).Caption = "Nombre :" & Val(v2) - ini4

Else

ItemNume = Trim(numcof(Num2).Caption)
ItemVale = Mid(valinv(Pic).Caption, 9, 50)
ItemDura = ItemDur(ItemNume).Dur

Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & ItemNume & SEP_CHAR & ItemVale & SEP_CHAR & ItemDura & SEP_CHAR & END_CHAR
Call SendData(Packet)

Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)

valcof(Num2).Caption = "Nombre :" & Val(v2) - ini4
End If

Call SendData("refresh" & SEP_CHAR & END_CHAR)
Call Form_Load
apr:
Call Form_Load

End Sub
Public Sub metrecofre_Click()
Dim Packet As String
Dim itemdurabi
If Pic = vbNullString Then variable = MsgBox("Aucun slot sélectioné!!!", vbCritical, "Erreur")
If Pic = vbNullString Then GoTo l:
If Num2 = vbNullString Then variable = MsgBox("Veuillez sélectioner un slot dans le coffre S.V.P!!", vbCritical, "Erreur")
If Num2 = vbNullString Then GoTo l:
v = Mid(valinv(Pic).Caption, 9, 100)
Va5 = Val(v)
Inum5 = GetPlayerInvItemNum(MyIndex, Pic)
If v = vbNullString And label1(Pic).Caption = vbNullString Then variable = MsgBox("Aucun objet dans le slot" & Pic & " de l'inventaire!!!", vbCritical, "Erreur")
If v = vbNullString And label1(Pic).Caption = vbNullString Then GoTo l:
If Label3(Num2).Caption = vbNullString Then
Else
If Item(Inum5).Type = 12 And Label3(Num2).Caption = label1(Pic).Caption Then
Else
variable = MsgBox("Il ya a déjà un objet dans le slot" & Num2 & " du coffre!!", vbCritical, "Erreur")
GoTo l:
End If
End If
If Num2 <= 0 Then variable = MsgBox("Veuillez séléctioner un slot dans le coffre S.V.P!!", vbCritical, "Erreur")
If Num2 <= 0 Then GoTo l:
ini = InputBox("Conbiens d'objet(s) voulez vous métre dans le coffre?", "Demande")
ini2 = Val(ini)
If ini2 > Va5 Then variable = MsgBox("Valeur supérieur au nombre d'objet!!", vbCritical, "Erreur")
If ini2 > Va5 Then GoTo l:
If ini = vbNullString Or ini = "0" Or ini = 0 Then GoTo l:
On Error Resume Next
valinv(Pic).Caption = "Nombre :" & GetPlayerInvItemValue(MyIndex, Pic) - ini
If Err.Number = 13 Then variable = MsgBox("Veuillez saisir un nombre svp!!!", vbCritical, "Erreur")
If Err.Number = 13 Then GoTo l:
If Item(Inum5).Type = 12 And Label3(Num2).Caption = label1(Pic).Caption Or Label3(Num2).Caption = vbNullString Then
itemdurabi = ItemDur(Inum5).Dur
Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & GetPlayerInvItemNum(MyIndex, Pic) & SEP_CHAR & Val(Mid(valcof(Num2).Caption, 9, 50)) + ini2 & SEP_CHAR & END_CHAR

Call SendData(Packet)

If Va5 - ini2 > 0 Then
Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & Inum5 & SEP_CHAR & Va5 - ini2 & SEP_CHAR & itemdurabi & SEP_CHAR & END_CHAR

Call SendData(Packet)
Else
Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)
End If

Else

If Item(Inum5).Type = 12 Then
Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & GetPlayerInvItemNum(MyIndex, Pic) & SEP_CHAR & ini & SEP_CHAR & END_CHAR

Call SendData(Packet)

If Va5 - ini2 > 0 Then
Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & Inum5 & SEP_CHAR & Va5 - ini2 & SEP_CHAR & ItemDur(Inum5).Dur & SEP_CHAR & END_CHAR

Call SendData(Packet)
Else
Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)
End If

Else

Packet = "COFFRE" & SEP_CHAR & Num2 & SEP_CHAR & GetPlayerInvItemNum(MyIndex, Pic) & SEP_CHAR & GetPlayerInvItemValue(MyIndex, Pic) & SEP_CHAR & END_CHAR

Call SendData(Packet)

Packet = "BANKITEM" & SEP_CHAR & Pic & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR

Call SendData(Packet)

End If
End If

Call SendData("refresh" & SEP_CHAR & END_CHAR)

Call Form_Load

Call Form_Load

l:

Call Form_Load
End Sub

