VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditArrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditeur de flêche"
   ClientHeight    =   3090
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   5054
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Configuration"
      TabPicture(0)   =   "frmEditArrows.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblRange"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblArrow"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdOk"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "scrlArrow"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "scrlRange"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtName"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txtName 
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
         TabIndex        =   9
         ToolTipText     =   "Nom de la flèche"
         Top             =   720
         Width           =   2775
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   240
         Max             =   30
         Min             =   1
         TabIndex        =   7
         Top             =   1560
         Value           =   1
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   5760
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   4
         ToolTipText     =   "Apparence de la flèche"
         Top             =   1080
         Width           =   540
         Begin VB.PictureBox picArrows 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   10
            ToolTipText     =   "Apparence de la flèche"
            Top             =   15
            Width           =   480
         End
      End
      Begin VB.HScrollBar scrlArrow 
         Height          =   255
         Left            =   3480
         Max             =   500
         Min             =   1
         TabIndex        =   3
         Top             =   720
         Value           =   1
         Width           =   2775
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Quitte la fenêtre d'édition et enregistre la flèche"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer la flèche"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nom:"
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
         TabIndex        =   8
         Top             =   480
         Width           =   345
      End
      Begin VB.Label lblArrow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Flêche:"
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
         Left            =   3480
         TabIndex        =   6
         ToolTipText     =   "Numéros de la flèche"
         Top             =   480
         Width           =   450
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         Caption         =   "Portée :"
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
         TabIndex        =   5
         ToolTipText     =   "Porté en cases de la flèche"
         Top             =   1320
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmEditArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Call ArrowEditorOk
End Sub

Private Sub Command1_Click()
    Call ArrowEditorCancel
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ArrowEditorCancel
End Sub

Private Sub scrlArrow_Change()
    lblArrow.Caption = "Flêche : " & scrlArrow.value
    Call AffSurfPic(DD_ArrowAnim, Me.picArrows, 3 * PIC_X, scrlArrow.value * PIC_Y)
    'picArrows.Top = (scrlArrow.value * 32) * -1
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = "Portée : " & scrlRange.value
End Sub

