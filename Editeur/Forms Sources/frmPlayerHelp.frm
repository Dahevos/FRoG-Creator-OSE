VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPlayerHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Panneau de commande"
   ClientHeight    =   3480
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3270
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlayerHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   5741
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   1879
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Commande"
      TabPicture(0)   =   "frmPlayerHelp.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Aide"
      TabPicture(1)   =   "frmPlayerHelp.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame8 
         Caption         =   "Conseil"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   2775
         Begin VB.TextBox Text1 
            Height          =   1575
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Text            =   "frmPlayerHelp.frx":0902
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Commande sans Cible"
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   2775
         Begin VB.CommandButton Command2 
            Caption         =   "Site Web du jeu"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   2535
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Demander de l'aide à un admin !"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   2535
         End
         Begin VB.CommandButton btnAbsent 
            Caption         =   "Information sur soi"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Commande avec Cible"
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         Begin VB.CommandButton Command1 
            Caption         =   "Informations sur le Joueur"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtPlayer 
            Height          =   240
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Nom de la Cible:"
            Height          =   135
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Panneau d'aide au Joueur"
         Height          =   135
         Index           =   1
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Panneau d'aide au Joueur"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmPlayerHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAbsent_Click()
Call SendGetStats
End Sub

Private Sub Command1_Click()
Call SendGetOtherStats(Trim$(txtPlayer.Text))
End Sub

Private Sub Command2_Click()
ShellExecute Me.hwnd, "open", ReadINI("CONFIG", "WebSite", (App.Path & "\Config\Client.ini")), vbNullString, App.Path, 1
End Sub

Private Sub Command3_Click()
Call SendGetAdminHelp
End Sub
