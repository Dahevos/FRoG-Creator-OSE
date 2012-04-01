VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{463051F7-93F6-433B-8C04-1B5EF7493179}#1.0#0"; "WinXPCEngine.ocx"
Begin VB.Form frmFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options des cartes par FTP"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar bar 
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton envcFTP 
      Caption         =   "Envoyer les cartes sur le FTP"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   3975
   End
   Begin VB.CheckBox actFTP 
      Caption         =   "Activer les cartes par FTP"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox url 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "http://"
      ToolTipText     =   "URL du FTP ex : http://frogcreator.leobaillard.org"
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton test 
      Caption         =   "Tester"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox rep 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Text            =   "/"
      ToolTipText     =   "Répertoir ou seront envoyer les cartes"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox hote 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Hote FTP ex: ftpperso.free.fr"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton sauv 
      Caption         =   "Sauvegarder"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   360
      Top             =   1200
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      ListBoxControl  =   0   'False
      PictureControl  =   0   'False
      FrameControl    =   0   'False
      OptionControl   =   0   'False
      ComboBoxControl =   0   'False
      DriveListBoxControl=   0   'False
      TabStripControl =   0   'False
      StatusBarControl=   0   'False
      SliderControl   =   0   'False
      ImageComboControl=   0   'False
      ListViewControl =   0   'False
      FileListBoxControl=   0   'False
      DirListBoxControl=   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "URL du FTP :"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   960
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Répertoir distant :"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   600
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hote FTP :"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmftp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub actFTP_Click()
    Dim i As Long
    If actFTP.value = Checked Then
        PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 1
        CarteFTP = True
        i = MsgBox("Voulez vous envoyer les cartes sur le FTP maitenant?", vbYesNo)
        If i = vbYes Then Call envcFTP_Click
    Else
        PutVar App.Path & "\Data.ini", "FTP", "ACTIF", 0
        CarteFTP = False
    End If
End Sub

Private Sub envcFTP_Click()
frmEnvFTP.Show vbModeless, frmftp
End Sub

Private Sub Form_Load()
    hote.text = GetVar(App.Path & "\Data.ini", "FTP", "HOTE")
    rep.text = GetVar(App.Path & "\Data.ini", "FTP", "REP")
    url.text = GetVar(App.Path & "\Data.ini", "FTP", "URL")
    actFTP.value = Val(GetVar(App.Path & "\Data.ini", "FTP", "ACTIF"))
    WindowsXPC1.InitSubClassing
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
frmCoFTP.Show vbModeless, frmftp
'bar.value = 0
'frmftp.bar.Visible = True
'Call TestConection(hote.text, nom.text, mdp.text, rep.text)
End Sub

Private Sub url_GotFocus()
url.SelStart = 0
url.SelLength = Len(url.text)
End Sub
