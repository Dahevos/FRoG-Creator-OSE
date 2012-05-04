VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmBClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloquer une Classe"
   ClientHeight    =   2985
   ClientLeft      =   1065
   ClientTop       =   2550
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
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
      TabCaption(0)   =   "Exception"
      TabPicture(0)   =   "frmBClass.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNum1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblNum2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblNum3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlNum1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOk"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlNum2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "scrlNum3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      Begin VB.HScrollBar scrlNum3 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   9
         Top             =   1800
         Width           =   4335
      End
      Begin VB.HScrollBar scrlNum2 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   8
         Top             =   1200
         Width           =   4335
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
         Top             =   2280
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
         Left            =   2520
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
      Begin VB.HScrollBar scrlNum1 
         Height          =   255
         Left            =   240
         Max             =   30
         TabIndex        =   1
         Top             =   600
         Width           =   4335
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
         Left            =   1200
         TabIndex        =   11
         Top             =   1560
         Width           =   75
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
         Left            =   1200
         TabIndex        =   10
         Top             =   960
         Width           =   75
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classe acceptée:"
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
         Left            =   105
         TabIndex        =   7
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classe acceptée:"
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
         Left            =   105
         TabIndex        =   6
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Classe acceptée:"
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
         Left            =   105
         TabIndex        =   5
         Top             =   360
         Width           =   1035
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
         Left            =   1200
         TabIndex        =   4
         Top             =   360
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmBClass"
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

Private Sub Form_Load()
    lblNum1.Caption = scrlNum1.value & " - " & Trim$(Class(scrlNum1.value).name)
    lblNum2.Caption = scrlNum2.value & " - " & Trim$(Class(scrlNum2.value).name)
    lblNum3.Caption = scrlNum3.value & " - " & Trim$(Class(scrlNum3.value).name)
    If EditorItemNum1 < scrlNum1.min Then EditorItemNum1 = scrlNum1.min
    scrlNum1.value = EditorItemNum1
    If EditorItemNum2 < scrlNum2.min Then EditorItemNum2 = scrlNum2.min
    scrlNum2.value = EditorItemNum2
    If EditorItemNum3 < scrlNum3.min Then EditorItemNum3 = scrlNum3.min
    scrlNum3.value = EditorItemNum3
End Sub

Private Sub scrlNum1_Change()
    lblNum1.Caption = scrlNum1.value & " - " & Trim$(Class(scrlNum1.value).name)
    EditorItemNum1 = CStr(scrlNum1.value)
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = scrlNum2.value & " - " & Trim$(Class(scrlNum2.value).name)
    EditorItemNum2 = CStr(scrlNum2.value)
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = scrlNum3.value & " - " & Trim$(Class(scrlNum3.value).name)
    EditorItemNum3 = CStr(scrlNum3.value)
End Sub

