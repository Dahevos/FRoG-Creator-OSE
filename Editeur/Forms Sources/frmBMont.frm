VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmBMont 
   Caption         =   "Bloquer une monture"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2990
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   388
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Info"
      TabPicture(0)   =   "frmBMont.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNum1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "scrlNum1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCancel"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum1 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bloquer si le joueur a l'objet monture equipper :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   270
         TabIndex        =   5
         Top             =   360
         Width           =   2940
      End
      Begin VB.Label lblNum1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3240
         TabIndex        =   4
         Top             =   360
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmBMont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

