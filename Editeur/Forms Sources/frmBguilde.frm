VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBguilde 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloquer des guildes"
   ClientHeight    =   1575
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
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
      TabPicture(0)   =   "frmBguilde.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdCancel"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOk"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "nom"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox nom 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         ToolTipText     =   "Entrez le nom de la guilde qui serat autorisée"
         Top             =   480
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
         TabIndex        =   2
         Top             =   960
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
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Autoriser la guilde :"
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
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmBguilde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub cmdOk_Click()
NomGuilde = nom.Text
Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub Form_Load()
nom.Text = NomGuilde
End Sub
