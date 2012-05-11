VERSION 5.00
Begin VB.Form frmCoFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connexion au serveur FTP"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox coauto 
      Caption         =   "Connexion automatique"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1065
      Width           =   3015
   End
   Begin VB.CommandButton annul 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2250
      TabIndex        =   6
      Top             =   1430
      Width           =   1695
   End
   Begin VB.TextBox mdp 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Votre mot de passe"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox nom 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Votre nom d'utilisateur"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton bt 
      Caption         =   "Connexion"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CheckBox memo 
      Caption         =   "Mémoriser le mot de passe"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mot de passe :"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nom d'utilisateur :"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1245
   End
End
Attribute VB_Name = "frmCoFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub annul_Click()
Unload Me
End Sub

Private Sub bt_Click()
If Memo.value = Checked Then
    WriteINI "FTP", "NOM", nom.Text, App.Path & "\Config.ini"
    WriteINI "FTP", "MDP", mdp.Text, App.Path & "\Config.ini"
    WriteINI "FTP", "AUTO", coauto.value, App.Path & "\Config.ini"
Else
    WriteINI "FTP", "NOM", vbNullString, App.Path & "\Config.ini"
    WriteINI "FTP", "MDP", vbNullString, App.Path & "\Config.ini"
    WriteINI "FTP", "AUTO", 0, App.Path & "\Config.ini"
End If
frmmsg.Show
Call Envoi(bt.Tag, nom.Text, mdp.Text, "Maps\map" & Player(MyIndex).Map & ".fcc", "map" & Player(MyIndex).Map & ".fcc", annul.Tag)
Call SendData("MAPDOWN" & END_CHAR)
Unload Me
End Sub

Private Sub Form_Load()
If ReadINI("FTP", "NOM", App.Path & "\Config.ini") <> vbNullString Then Memo.value = Checked
coauto.value = Val(ReadINI("FTP", "AUTO", App.Path & "\Config.ini"))
nom.Text = ReadINI("FTP", "NOM", App.Path & "\Config.ini")
mdp.Text = ReadINI("FTP", "MDP", App.Path & "\Config.ini")
End Sub

