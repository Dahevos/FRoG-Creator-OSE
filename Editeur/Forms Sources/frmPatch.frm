VERSION 5.00
Begin VB.Form frmPatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de patchs"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Cree 
      Caption         =   "Créer le patch"
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      ToolTipText     =   "Crèe le info.ini du patch"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.HScrollBar nbf 
      Height          =   255
      Left            =   120
      Max             =   3000
      Min             =   1
      TabIndex        =   3
      Top             =   1320
      Value           =   1
      Width           =   4215
   End
   Begin VB.TextBox cf 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      ToolTipText     =   "Chemins du fichier à remplacer (ex : pour le fichier sprites.png le chemin est ""GFX\"")"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox vf 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Nouvelle version du fichier à remplacer"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox nf 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      ToolTipText     =   "Nom du fichier à remplacer"
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox nexe 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "Nom de l'éxecutable du client (""Client"" par default)"
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox vp 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      ToolTipText     =   "La version du patch qui doit être supérieur aux versions des patchs précédent"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fichier numéros 1"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Chemins :"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Version du fichier :"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nom du fichier : "
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nom de l'exe du client :"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Version du patch :"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1290
   End
End
Attribute VB_Name = "frmPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MaxF As Long
Dim Fichier(1 To 3000) As Fichiers

Private Sub cf_Change()
Fichier(nbf.value).Chemins = Trim$(cf.Text)
End Sub

Private Sub Cree_Click()
Dim i As Long

If FileExist("\info.ini") Then Call Kill(App.Path & "\info.ini")

For i = 1 To 3000
    If Trim$(Fichier(i).nom) <= vbNullString Then MaxF = i - 1: Exit For
Next i

Call WriteINI("VERSION", "Version", Trim$(vp.Text), App.Path & "\info.ini")
Call WriteINI("VERSION", "GameFileName", Trim$(nexe.Text), App.Path & "\info.ini")
Call WriteINI("VERSION", "Max", Trim$(CStr(MaxF)), App.Path & "\info.ini")

For i = 1 To MaxF
    Call WriteINI("FILES", "FileName" & i, Fichier(i).nom, App.Path & "\info.ini")
    Call WriteINI("FILES", "FileVersion" & i, Fichier(i).version, App.Path & "\info.ini")
    Call WriteINI("FILES", "FilePath" & i, Fichier(i).Chemins, App.Path & "\info.ini")
Next i

Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
frmMirage.Show
End Sub

Private Sub nbf_Change()
Label6.Caption = "Fichier Numéro " & nbf.value

nf.Text = Fichier(nbf.value).nom
vf.Text = Fichier(nbf.value).version
cf.Text = Fichier(nbf.value).Chemins
End Sub

Private Sub nf_Change()
Fichier(nbf.value).nom = Trim$(nf.Text)
End Sub

Private Sub vf_Change()
Fichier(nbf.value).version = Trim$(vf.Text)
End Sub
