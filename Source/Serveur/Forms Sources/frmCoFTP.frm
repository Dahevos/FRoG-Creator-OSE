VERSION 5.00
Begin VB.Form frmCoFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connexion au serveur FTP"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton annul 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox memo 
      Caption         =   "Mémoriser le mot de passe"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton bt 
      Caption         =   "Connexion"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox nom 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      ToolTipText     =   "Votre nom d'utilisateur"
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox mdp 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Votre mot de passe"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nom d'utilisateur :"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Mot de passe :"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1050
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
If memo.value = Checked Then
    PutVar App.Path & "\Data.ini", "FTP", "NOM", nom.text
    PutVar App.Path & "\Data.ini", "FTP", "MDP", mdp.text
Else
    PutVar App.Path & "\Data.ini", "FTP", "NOM", vbNullString
    PutVar App.Path & "\Data.ini", "FTP", "MDP", vbNullString
End If
If bt.Caption = "Tester" Then
    Unload Me
    frmOptFTP.bar.value = 0
    frmOptFTP.bar.Visible = True
    Call TestConection(frmOptFTP.hote.text, nom.text, mdp.text, frmOptFTP.rep.text)
ElseIf bt.Caption = "Connexion" Then
    Unload Me
    Call frmEnvFTP.Env
End If
End Sub

Private Sub Form_Load()
If GetVar(App.Path & "\Data.ini", "FTP", "NOM") <> vbNullString Then memo.value = Checked
nom.text = GetVar(App.Path & "\Data.ini", "FTP", "NOM")
mdp.text = GetVar(App.Path & "\Data.ini", "FTP", "MDP")
End Sub
