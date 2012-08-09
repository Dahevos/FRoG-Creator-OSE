VERSION 5.00
Begin VB.Form frmclasseseditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Classes de FRoG Creator"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox lock 
      Caption         =   "Les joueurs ne peuvent pas sélectionner cette classe"
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   5880
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      Caption         =   "Apparence"
      Height          =   1095
      Left            =   120
      TabIndex        =   42
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   4680
      Width           =   7455
      Begin VB.HScrollBar scrlfem 
         Height          =   255
         Left            =   3720
         Max             =   1000
         TabIndex        =   46
         Top             =   600
         Width           =   3000
      End
      Begin VB.HScrollBar scrlhom 
         Height          =   255
         Left            =   200
         Max             =   1000
         TabIndex        =   45
         Top             =   600
         Width           =   3000
      End
      Begin VB.Label numsf 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   5520
         TabIndex        =   48
         Top             =   240
         Width           =   90
      End
      Begin VB.Label numsh 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2160
         TabIndex        =   47
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Apparence des Femmes :"
         Height          =   195
         Left            =   3720
         TabIndex        =   44
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Apparence des Hommes :"
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1830
      End
   End
   Begin VB.CommandButton save 
      Caption         =   "Sauvegarder"
      Height          =   300
      Left            =   3000
      TabIndex        =   20
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quand le joueur change de classe"
      Height          =   1815
      Left            =   3840
      TabIndex        =   30
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   2760
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
         Top             =   360
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
      Height          =   1575
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   2760
      Width           =   3615
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
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "N'entrez que des chiffres valide SVP!!!(sauf nom de classe)"
      Top             =   120
      Width           =   7455
      Begin VB.TextBox caske 
         Height          =   285
         Left            =   5520
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "Numeros du Casque"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox arme 
         Height          =   285
         Left            =   5520
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Numeros de l'arme"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox armure 
         Height          =   285
         Left            =   5520
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "Numeros de l'armure"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox bouclier 
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Text            =   "0"
         ToolTipText     =   "Numeros du Bouclier"
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
         ToolTipText     =   "Numéros de la carte"
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

Private Sub ajf_GotFocus()
ajf.SelStart = 0
ajf.SelLength = Len(ajf)
End Sub

Private Sub ajm_GotFocus()
ajm.SelStart = 0
ajm.SelLength = Len(ajm)
End Sub

Private Sub ajv_GotFocus()
ajv.SelStart = 0
ajv.SelLength = Len(ajv)
End Sub

Private Sub arme_GotFocus()
arme.SelStart = 0
arme.SelLength = Len(arme)
End Sub

Private Sub armure_GotFocus()
armure.SelStart = 0
armure.SelLength = Len(armure)
End Sub

Private Sub bouclier_GotFocus()
bouclier.SelStart = 0
bouclier.SelLength = Len(bouclier)
End Sub

Private Sub carted_GotFocus()
carted.SelStart = 0
carted.SelLength = Len(carted)
End Sub

Private Sub cartem_GotFocus()
cartem.SelStart = 0
cartem.SelLength = Len(cartem)
End Sub

Private Sub caske_GotFocus()
caske.SelStart = 0
caske.SelLength = Len(caske)
End Sub

Private Sub def_GotFocus()
def.SelStart = 0
def.SelLength = Len(def)
End Sub

Private Sub force_GotFocus()
force.SelStart = 0
force.SelLength = Len(force)
End Sub

Private Sub magi_GotFocus()
magi.SelStart = 0
magi.SelLength = Len(magi)
End Sub

Private Sub nom_GotFocus()
nom.SelStart = 0
nom.SelLength = Len(nom)
End Sub

Private Sub Save_Click()
On Error Resume Next
Call WriteINI("CLASS", "Name", frmclasseseditor.nom.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "MaleSprite", frmclasseseditor.scrlhom.value, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "FemaleSprite", frmclasseseditor.scrlfem.value, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "STR", Val(frmclasseseditor.force.text), App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "DEF", frmclasseseditor.def.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "SPEED", frmclasseseditor.vit.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "MAGI", frmclasseseditor.magi.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "MAP", frmclasseseditor.carted.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "X", Val(frmclasseseditor.xd.text), App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "Y", frmclasseseditor.yd.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("STARTUP", "Weapon", frmclasseseditor.arme.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("STARTUP", "Shield", frmclasseseditor.bouclier.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("STARTUP", "Armor", frmclasseseditor.armure.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("STARTUP", "Helmet", frmclasseseditor.caske.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASSCHANGE", "AddStr", frmclasseseditor.ajf.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASSCHANGE", "AddDef", Val(frmclasseseditor.ajd.text), App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASSCHANGE", "AddSpeed", frmclasseseditor.ajv.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASSCHANGE", "AddMagi", frmclasseseditor.ajm.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("DEATH", "Map", frmclasseseditor.cartem.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("DEATH", "x", frmclasseseditor.xm.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("DEATH", "y", frmclasseseditor.ym.text, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Call WriteINI("CLASS", "Locked", frmclasseseditor.lock.value, App.Path & "\Classes\Class" & frmclasseseditor.Tag & ".ini")
Me.Hide
End Sub

Private Sub scrlfem_Change()
    numsf.Caption = scrlfem.value
End Sub

Private Sub scrlhom_Change()
    numsh.Caption = scrlhom.value
End Sub

Private Sub vit_GotFocus()
vit.SelStart = 0
vit.SelLength = Len(vit)
End Sub

Private Sub xd_GotFocus()
xd.SelStart = 0
xd.SelLength = Len(xd)
End Sub

Private Sub xm_GotFocus()
xm.SelStart = 0
xm.SelLength = Len(xm)
End Sub

Private Sub yd_GotFocus()
yd.SelStart = 0
yd.SelLength = Len(yd)
End Sub

Private Sub ym_GotFocus()
ym.SelStart = 0
ym.SelLength = Len(ym)
End Sub
