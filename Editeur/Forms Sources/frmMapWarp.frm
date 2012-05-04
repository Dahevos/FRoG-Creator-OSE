VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMapWarp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Téléportation"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Téléportation à.."
      TabPicture(0)   =   "frmMapWarp.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblX"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblY"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlX"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlY"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancel"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMap"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "collco"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton collco 
         Caption         =   "Coller les coordonées"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         ToolTipText     =   "Coller les coordonées enregistrées précédement"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Définir..."
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
         Left            =   720
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtMap 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   5
         Text            =   "1"
         Top             =   480
         Width           =   3255
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
         Top             =   2040
         Width           =   2055
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
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   2
         Top             =   1320
         Width           =   3255
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   720
         Max             =   30
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblY 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   10
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblX 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Y"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "X"
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
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Carte :"
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
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmMapWarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    EditorWarpMap = Val(txtMap.Text)
    EditorWarpX = scrlX.value
    EditorWarpY = scrlY.value
    InDefTel = False
    Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub cmdCancel_Click()
    InDefTel = False
    Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub collco_Click()
scrlX.value = CoordX
scrlY.value = CoordY
txtMap.Text = CoordM
End Sub

Private Sub Command1_Click()
InDefTel = True
Me.Hide
frmMirage.SetFocus
End Sub

Private Sub form_Load()
    scrlX.Max = MAX_MAPX
    scrlY.Max = MAX_MAPY
    
    If EditorWarpX < scrlX.min Then EditorWarpX = scrlX.min
    scrlX.value = EditorWarpX
    If EditorWarpY < scrlY.min Then EditorWarpY = scrlY.min
    scrlY.value = EditorWarpY
    txtMap.Text = EditorWarpMap
End Sub

Private Sub scrlX_Change()
    lblX.Caption = CStr(scrlX.value)
End Sub

Private Sub scrlY_Change()
    lblY.Caption = CStr(scrlY.value)
End Sub

Private Sub txtMap_Change()
    If Val(txtMap.Text) <= 0 Or Val(txtMap.Text) > MAX_MAPS Then txtMap.Text = "1"
End Sub

