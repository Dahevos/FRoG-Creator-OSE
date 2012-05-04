VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmArena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attribut de l'Arène"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   ControlBox      =   0   'False
   Icon            =   "frmArena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   441
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Configuration du Spawn"
      TabPicture(0)   =   "frmArena.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNum3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNum2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblNum1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlNum3"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlNum2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOk"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCancel"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlNum1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.HScrollBar scrlNum1 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   5
         Top             =   600
         Width           =   4335
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
         Left            =   2520
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
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
         Left            =   360
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum2 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.HScrollBar scrlNum3 
         Height          =   255
         Left            =   2520
         Max             =   30
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
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
         Left            =   650
         TabIndex        =   11
         Top             =   360
         Width           =   75
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Map:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "X:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Y:"
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
         Left            =   2520
         TabIndex        =   8
         Top             =   960
         Width           =   135
      End
      Begin VB.Label lblNum2 
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
         Left            =   480
         TabIndex        =   7
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lblNum3 
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
         Left            =   2760
         TabIndex        =   6
         Top             =   960
         Width           =   1755
      End
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    scrlNum1.value = 0
    scrlNum2.value = 0
    scrlNum3.value = 0
    Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub cmdOk_Click()
    Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub form_Load()
    scrlNum2.Max = MAX_MAPX
    scrlNum3.Max = MAX_MAPY
    If Arena1 < scrlNum1.min Then Arena1 = scrlNum1.min
    scrlNum1.value = Arena1
    If Arena2 < scrlNum2.min Then Arena2 = scrlNum2.min
    scrlNum2.value = Arena2
    If Arena3 < scrlNum3.min Then Arena3 = scrlNum3.min
    scrlNum3.value = Arena3
End Sub

Private Sub scrlNum1_Change()
    lblNum1.Caption = scrlNum1.value
    Arena1 = CStr(scrlNum1.value)
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = scrlNum2.value
    Arena2 = CStr(scrlNum2.value)
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = scrlNum3.value
    Arena3 = CStr(scrlNum3.value)
End Sub
