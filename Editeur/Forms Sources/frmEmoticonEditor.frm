VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmoticonEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer un Emoticone"
   ClientHeight    =   3525
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   108
      Left            =   1200
      Top             =   2400
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6059
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
      TabCaption(0)   =   "Emoticone"
      TabPicture(0)   =   "frmEmoticonEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEmoticon"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Picture1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "scrlEmoticon"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCommand"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
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
         Left            =   2640
         TabIndex        =   8
         ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer l'émoticon"
         Top             =   2880
         Width           =   1215
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
         Left            =   840
         TabIndex        =   7
         ToolTipText     =   "Quitte la fenêtre d'édition et enregistre l'émoticon"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCommand 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "/"
         ToolTipText     =   "Commande que devront tapé les jouer pour faire l'émoticon"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.HScrollBar scrlEmoticon 
         Height          =   255
         Left            =   240
         Max             =   1000
         TabIndex        =   3
         Top             =   840
         Value           =   1
         Width           =   2775
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   3240
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   1
         ToolTipText     =   "Apparence de l'émoticon"
         Top             =   840
         Width           =   540
         Begin VB.PictureBox picEmoticons 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   15
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   9
            ToolTipText     =   "Apparence de l'émoticon"
            Top             =   15
            Width           =   480
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Commande :"
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
         Top             =   1440
         Width           =   810
      End
      Begin VB.Label lblEmoticon 
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
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Numéros de l'émoticon"
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label5 
         Caption         =   "Emoticone :"
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
         TabIndex        =   2
         ToolTipText     =   "Numéros de l'émoticon"
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEmoticonEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FA As Byte

Private Sub cmdOk_Click()
Dim i As Long
    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) = Trim$(txtCommand.Text) And i <> EditorIndex - 1 And Trim$(txtCommand.Text) <> "/" Then MsgBox "Cette commande est déjà utilisée.": Exit Sub
    Next i
    Call EmoticonEditorOk
End Sub

Private Sub Command1_Click()
    Call EmoticonEditorCancel
End Sub

Private Sub Form_Load()
'    picEmoticons.Top = (scrlEmoticon.value * 32) * -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EmoticonEditorCancel
End Sub

Private Sub scrlEmoticon_Change()
    'picEmoticons.Top = (scrlEmoticon.value * 32) * -1
    lblEmoticon.Caption = scrlEmoticon.value
End Sub

Private Sub Timer1_Timer()
FA = FA + 1
If FA > 12 Then FA = 1
Call AffSurfPic(DD_EmoticonSurf, picEmoticons, FA * PIC_X, scrlEmoticon.value * PIC_Y)
 '   If picEmoticons.Left < -(10 * 32) Then picEmoticons.Left = 0
  '  picEmoticons.Left = picEmoticons.Left - 32
End Sub

Private Sub txtCommand_Change()
Dim i As String
i = txtCommand.Text
    If Mid$(i, 1, 1) <> "/" Then
        If Trim$(i) = vbNullString Then txtCommand.Text = "/": Exit Sub
        txtCommand.Text = "/" & i
        txtCommand.SelStart = Len(txtCommand.Text)
    End If
End Sub
