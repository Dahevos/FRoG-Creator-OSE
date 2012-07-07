VERSION 5.00
Begin VB.Form frmclasseseditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Classes de FRoG Creator"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox lock 
      Caption         =   "Les joueurs ne peuvent pas sélectionner cette classe"
      Height          =   255
      Left            =   240
      TabIndex        =   53
      Top             =   7440
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apparence"
      Height          =   2415
      Left            =   120
      TabIndex        =   42
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   4920
      Width           =   7455
      Begin VB.HScrollBar scrlfem 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   50
         Top             =   600
         Width           =   3000
      End
      Begin VB.HScrollBar scrlhom 
         Height          =   255
         Left            =   180
         Max             =   1000
         TabIndex        =   49
         Top             =   600
         Width           =   3000
      End
      Begin VB.PictureBox fem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   3720
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   45
         Top             =   1080
         Width           =   1020
         Begin VB.PictureBox femme 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000007&
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   15
            ScaleHeight     =   62
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   62
            TabIndex        =   46
            Top             =   15
            Width           =   960
         End
      End
      Begin VB.PictureBox mal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1020
         Left            =   240
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   43
         Top             =   1080
         Width           =   1020
         Begin VB.PictureBox homme 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000007&
            ForeColor       =   &H80000008&
            Height          =   960
            Left            =   15
            ScaleHeight     =   62
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   62
            TabIndex        =   44
            Top             =   15
            Width           =   960
         End
      End
      Begin VB.Label numsf 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   5520
         TabIndex        =   52
         Top             =   240
         Width           =   90
      End
      Begin VB.Label numsh 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2160
         TabIndex        =   51
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Apparence des Femmes :"
         Height          =   195
         Left            =   3720
         TabIndex        =   48
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Apparence des Hommes :"
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.CommandButton save 
      Caption         =   "Sauvegarder"
      Height          =   300
      Left            =   3000
      TabIndex        =   20
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quand le joueur change de classe"
      Height          =   1815
      Left            =   3840
      TabIndex        =   30
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   3000
      Width           =   3735
      Begin VB.TextBox ajv 
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Text            =   "0"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox ajd 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox ajf 
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         Text            =   "0"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox ajm 
         Height          =   285
         Left            =   1920
         TabIndex        =   19
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Ajout de Magie :"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Ajout de Vitesse :"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Ajout de Défense :"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Ajout de Force :"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A la mort du joueur"
      Height          =   1815
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   3000
      Width           =   3615
      Begin VB.CommandButton collco 
         Caption         =   "Coller les coordonées"
         Height          =   255
         Left            =   960
         TabIndex        =   54
         ToolTipText     =   "Coller les coordonées enregistrées précédement"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox ym 
         Height          =   285
         Left            =   1920
         TabIndex        =   15
         Text            =   "0"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox cartem 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Text            =   "0"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox xm 
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Carte de réapparition :"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1560
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Pos.X de réapparition :"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1605
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Pos.Y de réapparition :"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Général"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   120
      Width           =   7455
      Begin VB.CommandButton collco2 
         Caption         =   "Coller les coordonées"
         Height          =   255
         Left            =   2880
         TabIndex        =   55
         ToolTipText     =   "Coller les coordonées enregistrées précédement"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox caske 
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "Numero du Casque"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox arme 
         Height          =   285
         Left            =   5520
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Numero de l'arme"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox armure 
         Height          =   285
         Left            =   5520
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "Numero de l'armure"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox bouclier 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Numero du Bouclier"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox yd 
         Height          =   285
         Left            =   5520
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Nombre de Y(30max)"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox carted 
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Numéro de la carte"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox xd 
         Height          =   285
         Left            =   5520
         TabIndex        =   7
         Text            =   "0"
         ToolTipText     =   "Nombre de X (30max)"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox magi 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Nombre de Magie"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox force 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Nombre de Force"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox def 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Nombre de Défense"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox vit 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Nombre de Vitesse"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox nom 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "nomclasse"
         ToolTipText     =   "Nom de la classe"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Armure de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   41
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Casque de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   40
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Arme de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   39
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Bouclier de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   38
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Carte de départ :"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Position en X de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Position en Y de départ :"
         Height          =   195
         Left            =   3720
         TabIndex        =   26
         Top             =   600
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nom de la Classe :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Force de départ :"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Défense de départ :"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vitesse de départ :"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Magie de départ :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmclasseseditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ajd_GotFocus()
ajd.SelStart = 0
ajd.SelLength = Len(ajd)
End Sub
Private Sub ajd_change()
If Not IsNumeric(ajd.Text) Then
ajd.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub ajf_GotFocus()
ajf.SelStart = 0
ajf.SelLength = Len(ajf)
End Sub
Private Sub ajf_change()
If Not IsNumeric(ajf.Text) Then
ajf.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub ajm_GotFocus()
ajm.SelStart = 0
ajm.SelLength = Len(ajm)
End Sub
Private Sub ajm_change()
If Not IsNumeric(ajm.Text) Then
ajm.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub ajv_GotFocus()
ajv.SelStart = 0
ajv.SelLength = Len(ajv)
End Sub
Private Sub ajv_change()
If Not IsNumeric(ajv.Text) Then
ajv.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub arme_GotFocus()
arme.SelStart = 0
arme.SelLength = Len(arme)
End Sub
Private Sub arme_change()
If Not IsNumeric(arme.Text) Then
arme.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub armure_GotFocus()
armure.SelStart = 0
armure.SelLength = Len(armure)
End Sub
Private Sub armure_change()
If Not IsNumeric(armure.Text) Then
armure.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub bouclier_GotFocus()
bouclier.SelStart = 0
bouclier.SelLength = Len(bouclier)
End Sub
Private Sub bouclier_change()
If Not IsNumeric(bouclier.Text) Then
bouclier.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub carted_GotFocus()
carted.SelStart = 0
carted.SelLength = Len(carted)
End Sub
Private Sub carted_change()
If Not IsNumeric(carted.Text) Then
carted.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub cartem_GotFocus()
cartem.SelStart = 0
cartem.SelLength = Len(cartem)
End Sub
Private Sub cartem_change()
If Not IsNumeric(cartem.Text) Then
cartem.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub caske_GotFocus()
caske.SelStart = 0
caske.SelLength = Len(caske)
End Sub

Private Sub collco_Click()
cartem.Text = CoordM
xm.Text = CoordX
ym.Text = CoordY
End Sub

Private Sub collco2_Click()
carted.Text = CoordM
xd.Text = CoordX
yd.Text = CoordY
End Sub

Private Sub def_GotFocus()
def.SelStart = 0
def.SelLength = Len(def)
End Sub
Private Sub def_change()
If Not IsNumeric(def.Text) Then
def.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub force_change()
If Not IsNumeric(force.Text) Then
force.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub force_GotFocus()
force.SelStart = 0
force.SelLength = Len(force)
End Sub

Private Sub Form_Load()
scrlhom.Max = MAX_DX_SPRITE
scrlfem.Max = MAX_DX_SPRITE
End Sub

Private Sub Form_Resize()
Exit Sub
End Sub

Private Sub magi_GotFocus()
magi.SelStart = 0
magi.SelLength = Len(magi)
End Sub
Private Sub magi_change()
If Not IsNumeric(magi.Text) Then
magi.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub nom_GotFocus()
nom.SelStart = 0
nom.SelLength = Len(nom)
End Sub

Private Sub Save_Click()

On Error GoTo ereu:
Dim PathServ As String

PathServ = Mid$(App.Path, 1, Len(App.Path) - Len(Dir$(App.Path, vbDirectory))) & "Serveur"

frmmsg.Show
If HORS_LIGNE = 1 Then GoTo hl:
If LCase$(Dir$(PathServ, vbDirectory)) <> "serveur" Then Call MsgBox("Dossier du serveur introuvable les modifications niveau serveur ne seront pas prises en comptes."): GoTo hl:

Call WriteINI("CLASS", "Name", frmclasseseditor.nom.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MaleSprite", frmclasseseditor.scrlhom.value, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "FemaleSprite", frmclasseseditor.scrlfem.value, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "STR", Val(frmclasseseditor.force.Text), PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "DEF", frmclasseseditor.def.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "SPEED", frmclasseseditor.vit.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MAGI", frmclasseseditor.magi.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MAP", frmclasseseditor.carted.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "X", Val(frmclasseseditor.xd.Text), PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "Y", frmclasseseditor.yd.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Weapon", frmclasseseditor.arme.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Shield", frmclasseseditor.bouclier.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Armor", frmclasseseditor.armure.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Helmet", frmclasseseditor.caske.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddStr", frmclasseseditor.ajf.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddDef", Val(frmclasseseditor.ajd.Text), PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddSpeed", frmclasseseditor.ajv.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddMagi", frmclasseseditor.ajm.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "Map", frmclasseseditor.cartem.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "x", frmclasseseditor.xm.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "y", frmclasseseditor.ym.Text, PathServ & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "Locked", frmclasseseditor.lock.value, PathServ & "\Classes\Class" & classe & ".ini")
hl:
Call WriteINI("CLASS", "Name", frmclasseseditor.nom.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MaleSprite", frmclasseseditor.scrlhom.value, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "FemaleSprite", frmclasseseditor.scrlfem.value, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "STR", Val(frmclasseseditor.force.Text), App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "DEF", frmclasseseditor.def.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "SPEED", frmclasseseditor.vit.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MAGI", frmclasseseditor.magi.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "MAP", frmclasseseditor.carted.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "X", Val(frmclasseseditor.xd.Text), App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "Y", frmclasseseditor.yd.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Weapon", frmclasseseditor.arme.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Shield", frmclasseseditor.bouclier.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Armor", frmclasseseditor.armure.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("STARTUP", "Helmet", frmclasseseditor.caske.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddStr", frmclasseseditor.ajf.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddDef", Val(frmclasseseditor.ajd.Text), App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddSpeed", frmclasseseditor.ajv.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASSCHANGE", "AddMagi", frmclasseseditor.ajm.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "Map", frmclasseseditor.cartem.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "x", frmclasseseditor.xm.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("DEATH", "y", frmclasseseditor.ym.Text, App.Path & "\Classes\Class" & classe & ".ini")
Call WriteINI("CLASS", "Locked", frmclasseseditor.lock.value, App.Path & "\Classes\Class" & classe & ".ini")
Call Unload(frmmsg)
Call ChargerClasses
Me.Hide
Exit Sub
ereu:
Call MsgBox("N'entrez que des chiffres s'il vous plait.")
Call MsgBox("Erreur rencontrez : " & Err.Number & " : " & Err.description, vbCritical)
Call EcrireEtat(Err.Number & " " & Err.description & " N'entrez que des chiffres s'il vous plait.")
End Sub

Private Sub scrlfem_Change()
On Error Resume Next
    numsf.Caption = scrlfem.value
    Call PrepareSprite(scrlfem.value)
    Call AffSurfPic(DD_SpriteSurf(scrlfem.value), femme, 0, 0)
End Sub

Private Sub scrlhom_Change()
On Error Resume Next
    numsh.Caption = scrlhom.value
    Call PrepareSprite(scrlhom.value)
    Call AffSurfPic(DD_SpriteSurf(scrlhom.value), homme, 0, 0)
End Sub

Private Sub vit_GotFocus()
vit.SelStart = 0
vit.SelLength = Len(vit)
End Sub
Private Sub vit_change()
If Not IsNumeric(vit.Text) Then
vit.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub xd_GotFocus()
xd.SelStart = 0
xd.SelLength = Len(xd)
End Sub
Private Sub xd_change()
If Not IsNumeric(xd.Text) Then
xd.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub xm_GotFocus()
xm.SelStart = 0
xm.SelLength = Len(xm)
End Sub
Private Sub xm_change()
If Not IsNumeric(xm.Text) Then
xm.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub yd_GotFocus()
yd.SelStart = 0
yd.SelLength = Len(yd)
End Sub
Private Sub yd_change()
If Not IsNumeric(yd.Text) Then
yd.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
Private Sub ym_GotFocus()
ym.SelStart = 0
ym.SelLength = Len(ym)
End Sub
Private Sub ym_change()
If Not IsNumeric(ym.Text) Then
ym.Text = "0"
MsgBox ("Veullez remplir cette case par une valeur numérique")
End If
End Sub
