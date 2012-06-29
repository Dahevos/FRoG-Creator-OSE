VERSION 5.00
Begin VB.Form frmEnvFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envoyer les cartes sur un FTP"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin Serveur.ctlProgressBar bar 
      Height          =   255
      Left            =   120
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      Appearance      =   1
   End
   Begin VB.CommandButton annul 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton envoyer 
      Caption         =   "Envoyer"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox fin 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "10"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox debut 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label temps 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label etat 
      Alignment       =   2  'Center
      Caption         =   " "
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Carte :"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   270
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "à"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   270
      Width           =   90
   End
End
Attribute VB_Name = "frmEnvFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExStop As Boolean

Private Sub annul_Click()
Unload Me
End Sub

Private Sub debut_Change()
If debut.text = vbNullString Then debut.text = "1"
If Not IsNumeric(debut.text) Then MsgBox "Entrez seulement des chiffres s'il vous plait.", vbCritical: debut.text = "1": Exit Sub
If Val(debut.text) <= 0 Then MsgBox "Entrez une valeur supérieur à zéro", vbCritical: debut.text = "1": Exit Sub
If Val(debut.text) > MAX_MAPS Then MsgBox "Entrez une valeur inférieur au maximum de cartes", vbCritical: debut.text = "1": Exit Sub
End Sub

Private Sub debut_GotFocus()
debut.SelStart = 0
debut.SelLength = Len(debut.text)
End Sub

Private Sub envoyer_Click()
frmCoFTP.bt.Caption = "Connexion"
frmCoFTP.Show vbModeless, frmEnvFTP
End Sub

Private Sub fin_Change()
If fin.text = vbNullString Then fin.text = MAX_MAPS
If Not IsNumeric(fin.text) Then MsgBox "Entrez seulement des chiffres s'il vous plait.", vbCritical: fin.text = MAX_MAPS: Exit Sub
If Val(fin.text) <= 0 Then MsgBox "Entrez une valeur supérieur à zéro", vbCritical: fin.text = MAX_MAPS: Exit Sub
If Val(fin.text) > MAX_MAPS Then MsgBox "Entrez une valeur inférieur au maximum de cartes", vbCritical: fin.text = MAX_MAPS: Exit Sub
End Sub

Private Sub fin_GotFocus()
fin.SelStart = 0
fin.SelLength = Len(fin.text)
End Sub

Private Sub Form_Load()
ExStop = False
fin.text = MAX_MAPS
End Sub

Private Sub Form_Unload(Cancel As Integer)
ExStop = True
End Sub

Public Sub Env()
Dim Connex As Long
Dim i As Long
Dim t As Long
On Error GoTo er:
Connex = 0

If Val(debut.text) > Val(fin.text) Then MsgBox "La 1er valeur ne peut pas être supérieur à la 2éme", vbCritical: Exit Sub

etat.Enabled = True
temps.Enabled = True
debut.Enabled = False
fin.Enabled = False
envoyer.Enabled = False
annul.Enabled = False

bar.value = 0
etat.Caption = "Connexion..."
If ExStop Then Call FermerFTP(Connex): GoTo fin:
NewDoEvents
Connex = ConnexionFTP(frmOptFTP.hote.text, frmCoFTP.nom.text, frmCoFTP.mdp.text)
If Connex = 0 Then GoTo fin:
bar.Max = Val(fin.text) + 10
If ExStop Then Call FermerFTP(Connex): GoTo fin:

For i = Val(debut.text) To Val(fin.text)
    bar.value = i
    etat.Caption = "Envoie de la carte" & i
    If ExStop Then Call FermerFTP(Connex): GoTo fin:
    NewDoEvents
    If i = Val(debut.text) Then t = Timer
    Call EnvoiFTP(Connex, frmOptFTP.hote.text, frmCoFTP.nom.text, frmCoFTP.mdp.text, "maps\map" & i & ".fcc", "map" & i & ".fcc", frmOptFTP.rep)
    If i = Val(debut.text) Then t = Timer - t: temps.Caption = "Temps théorique : " & t * (Val(fin.text) - i) & "s" Else temps.Caption = "Temps théorique : " & t * (Val(fin.text) - i) & "s"
    If ExStop Then Call FermerFTP(Connex): GoTo fin:
Next i

bar.value = bar.Max - 2
etat.Caption = "Fermeture de la connexion..."
If ExStop Then Call FermerFTP(Connex): GoTo fin:
Call FermerFTP(Connex)
bar.value = bar.Max

MsgBox "Envoie terminé.", vbInformation
etat.Enabled = False
temps.Enabled = False
etat.Caption = vbNullString
bar.value = bar.Min
temps.Caption = vbNullString
debut.Enabled = True
fin.Enabled = True
envoyer.Enabled = True
annul.Enabled = True

Exit Sub
er:
MsgBox "Erreur pendant l'envoie des cartes", vbCritical
fin:
etat.Enabled = False
bar.State = ccStateError
temps.Enabled = False
etat.Caption = vbNullString
bar.value = bar.Min
temps.Caption = vbNullString
debut.Enabled = True
fin.Enabled = True
envoyer.Enabled = True
annul.Enabled = True
bar.State = ccStateError
If Connex <> 0 Then Call FermerFTP(Connex)
End Sub
