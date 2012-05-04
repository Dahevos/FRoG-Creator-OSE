VERSION 5.00
Begin VB.Form frmOptFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option des cartes par FTP"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin Serveur.ctlProgressBar bar 
      Height          =   255
      Left            =   840
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      Appearance      =   1
   End
   Begin VB.CommandButton sauv 
      Caption         =   "Sauvegarder"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox hote 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Hote FTP ex: ftpperso.free.fr"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox rep 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Text            =   "/"
      ToolTipText     =   "Répertoir ou seront envoyer les cartes"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton test 
      Caption         =   "Tester"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "http://"
      ToolTipText     =   "URL du FTP ex : http://frogcreator.leobaillard.org"
      Top             =   840
      Width           =   1935
   End
   Begin VB.CheckBox actFTP 
      Caption         =   "Activer les cartes par FTP"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.CommandButton envcFTP 
      Caption         =   "Envoyer les cartes sur le FTP"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hote FTP :"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Répertoire distant :"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "URL du FTP :"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   990
   End
End
Attribute VB_Name = "frmOptFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub actFTP_Click()
    Dim i As Long
    If actFTP.value = Checked Then
        If Val(GetVar(App.Path & "\Data.ini", "FTP", "ACTIF")) = 0 Then
            PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 1
            CarteFTP = True
            i = MsgBox("Voulez vous envoyer les cartes sur le FTP maitenant?", vbYesNo)
            If i = vbYes Then Call envcFTP_Click
        End If
        PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 1
        CarteFTP = True
    Else
        PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 0
        CarteFTP = False
    End If
End Sub

Private Sub envcFTP_Click()
frmEnvFTP.Show vbModeless, frmOptFTP
End Sub

Private Sub Form_Load()
    hote.text = GetVar(App.Path & "\Data.ini", "FTP", "HOTE")
    rep.text = GetVar(App.Path & "\Data.ini", "FTP", "REP")
    url.text = GetVar(App.Path & "\Data.ini", "FTP", "URL")
    actFTP.value = Val(GetVar(App.Path & "\Data.ini", "FTP", "ACTIF"))
End Sub

Private Sub hote_GotFocus()
hote.SelStart = 0
hote.SelLength = Len(hote.text)
End Sub

Private Sub rep_GotFocus()
rep.SelStart = 0
rep.SelLength = Len(rep.text)
End Sub

Private Sub sauv_Click()
    PutVar App.Path & "\Data.ini", "FTP", "HOTE", hote.text
    PutVar App.Path & "\Data.ini", "FTP", "REP", rep.text
    PutVar App.Path & "\Data.ini", "FTP", "URL", url.text
    If actFTP.value = Checked Then PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 1 Else PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 0
    Unload Me
End Sub

Private Sub test_Click()
frmCoFTP.bt.Caption = "Tester"
frmCoFTP.Show vbModeless, frmOptFTP
'bar.value = 0
'frmoptftp.bar.Visible = True
'Call TestConection(hote.text, nom.text, mdp.text, rep.text)
End Sub

Private Sub url_GotFocus()
url.SelStart = 0
url.SelLength = Len(url.text)
End Sub

