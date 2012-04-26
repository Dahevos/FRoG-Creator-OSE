VERSION 5.00
Begin VB.Form frmCpnjmouv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editer les mouvement du PNJ de la Carte"
   ClientHeight    =   4875
   ClientLeft      =   180
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   363
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Direction du PNJ"
      Height          =   2355
      Left            =   2460
      TabIndex        =   27
      ToolTipText     =   "Direction du PNJ si il est immobile"
      Top             =   0
      Width           =   2895
      Begin VB.VScrollBar scrlSkinY 
         Height          =   1755
         Left            =   2520
         Max             =   1
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
      Begin VB.HScrollBar scrlSkinX 
         Height          =   255
         Left            =   960
         Max             =   1
         TabIndex        =   34
         Top             =   2040
         Width           =   1515
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   960
         ScaleHeight     =   115
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   99
         TabIndex        =   32
         Top             =   240
         Width           =   1515
         Begin VB.PictureBox skin 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   840
            Left            =   0
            ScaleHeight     =   56
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   36
            TabIndex        =   33
            Top             =   0
            Width           =   540
         End
      End
      Begin VB.CommandButton droite 
         Caption         =   "Droite"
         Height          =   240
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   780
      End
      Begin VB.CommandButton gauche 
         Caption         =   "Gauche"
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   540
         Width           =   780
      End
      Begin VB.CommandButton bas 
         Caption         =   "Bas"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1140
         Width           =   780
      End
      Begin VB.CommandButton haut 
         Caption         =   "Haut"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.CheckBox imob 
      Caption         =   "Immobile"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      ToolTipText     =   "Le pnj ne bougera pas si cette option est cochée"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CheckBox th 
      Caption         =   "Tout au hasard"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Toutes les Coordonées seront définit au hasard"
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mouvements du PNJ"
      Height          =   2055
      Left            =   540
      TabIndex        =   16
      Top             =   2460
      Width           =   4575
      Begin VB.CommandButton collco 
         Caption         =   "Coller..."
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   26
         ToolTipText     =   "Coller les coordonées"
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton collco 
         Caption         =   "Coller..."
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   25
         ToolTipText     =   "Coller les coordonées"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton defxy2 
         Caption         =   "Définir..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         ToolTipText     =   "Définir les Coordonées sur la carte"
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton defxy1 
         Caption         =   "Définir..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3360
         TabIndex        =   12
         ToolTipText     =   "Définir les Coordonées sur la carte"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox boucle 
         Caption         =   "Le PNJ fait une Ronde"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "Le PNJ fera le chemin de façon à faire une boucle c'est-à-dire tourner autour d'un block par exemple "
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox mvh 
         Caption         =   "Au hasard"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         ToolTipText     =   "Les Coordonées seront définit au hasard"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox x 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "Coordonner X ou va se diriger le PNJ"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox y 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "Coordonner Y ou va se diriger le PNJ"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox y 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "Coordonner Y ou va se diriger le PNJ"
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox x 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   8
         Text            =   "0"
         ToolTipText     =   "Coordonner X vers ou va se diriger le PNJ"
         Top             =   570
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Direction du 2e mouvement en Y :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Direction du 2e mouvement en X :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Direction du 1er mouvement en Y :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   960
         Width           =   2460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Direction du 1er mouvement en X :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   600
         Width           =   2460
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Position de départ du PNJ"
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton collco 
         Caption         =   "Coller..."
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         ToolTipText     =   "Coller les coordonées"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton defxy 
         Caption         =   "Définir..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Définir les Coordonées sur la carte"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox psh 
         Caption         =   "Au hasard"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Les Coordonées seront définit au hasard"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox x 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Coordonner X"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox y 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   840
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "Coordonner Y"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X :"
         Enabled         =   0   'False
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   600
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmCpnjmouv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim DIRC As Byte
Dim PNJNum As Long

Private Sub bas_Click()
Dim sprt As Long
On Error Resume Next
sprt = Npc(PNJNum).sprite
DIRC = DIR_DOWN
If PNJNum > 0 And PNJNum < MAX_NPCS Then
Call PrepareSprite(sprt)
Call AffSurfPic(DD_SpriteSurf(sprt), frmCpnjmouv.skin, 0, DIR_DOWN * (DDSD_Character(sprt).lHeight / 4))
End If
End Sub


Private Sub collco_Click(Index As Integer)
x(Index).Text = CoordX
y(Index).Text = CoordY
End Sub

Private Sub Command2_Click()
Unload Me
InMouvEditor = False
frmMapProperties.Show
End Sub

Private Sub defxy_Click()
frmMapProperties.Hide
frmMirage.SetFocus
cordo = 0
End Sub

Private Sub defxy1_Click()
frmMapProperties.Hide
frmMirage.SetFocus
cordo = 1
End Sub

Private Sub defxy2_Click()
frmMapProperties.Hide
frmMirage.SetFocus
cordo = 2
End Sub

Private Sub droite_Click()
Dim sprt As Long
On Error Resume Next
sprt = Npc(PNJNum).sprite
DIRC = DIR_RIGHT
If PNJNum > 0 And PNJNum < MAX_NPCS Then
Call PrepareSprite(sprt)
Call AffSurfPic(DD_SpriteSurf(sprt), frmCpnjmouv.skin, 0, DIR_RIGHT * (DDSD_Character(sprt).lHeight / 4))
End If
End Sub

Private Sub Form_Load()
Dim sprt As Long
Dim sd As Byte
x(0).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x
y(0).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y
x(1).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x1
y(1).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y1
x(2).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x2
y(2).Text = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y2
mvh.value = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Hasardm
psh.value = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Hasardp
If Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Imobile > 0 Then imob.value = Checked: gauche.Enabled = True: droite.Enabled = True: haut.Enabled = True: bas.Enabled = True: skin.Enabled = True Else imob.value = Unchecked: gauche.Enabled = False: droite.Enabled = False: haut.Enabled = False: bas.Enabled = False: skin.Enabled = False
boucle.value = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).boucle
If mvh.value = Checked And psh.value = Checked Then th.value = Checked
If Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Imobile > 0 Then
    DIRC = Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Imobile - 1
Else
    DIRC = DIR_DOWN
End If
PNJNum = frmMapProperties.cmbNpc(EditorMouvIndex - 1).ListIndex
On Error Resume Next
sprt = Npc(PNJNum).sprite
Call PrepareSprite(sprt)
skin.height = (DDSD_Character(sprt).lHeight / 4) '* Screen.TwipsPerPixelY
skin.Width = (DDSD_Character(sprt).lWidth / 4) '* Screen.TwipsPerPixelX

If skin.height > Picture1.height Then
    scrlSkinX.Enabled = True
    scrlSkinX.Max = Picture1.height - skin.height
Else
    scrlSkinX.Enabled = False
End If

If skin.Width > Picture1.Width Then
    scrlSkinY.Enabled = True
    scrlSkinY.Max = Picture1.Width - skin.Width
Else
    scrlSkinY.Enabled = False
End If

If PNJNum > 0 And PNJNum < MAX_NPCS Then
Call AffSurfPic(DD_SpriteSurf(sprt), frmCpnjmouv.skin, 0, DIRC * (DDSD_Character(sprt).lHeight / 4))
End If

If mvh.value = Checked Then
    boucle.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    For i = 1 To 2
        x(i).Enabled = False
        y(i).Enabled = False
    Next i
    defxy1.Enabled = False
    defxy2.Enabled = False
    collco(1).Enabled = False
    collco(2).Enabled = False
End If
If psh.value = Checked Then
    Label1.Enabled = False
    Label2.Enabled = False
    x(0).Enabled = False
    y(0).Enabled = False
    defxy.Enabled = False
    collco(0).Enabled = False
End If
End Sub

Private Sub Form_Terminate()
frmMapProperties.Show
InMouvEditor = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMapProperties.Show
InMouvEditor = False
End Sub

Private Sub gauche_Click()
Dim sprt As Long
On Error Resume Next
sprt = Npc(PNJNum).sprite
DIRC = DIR_LEFT
If PNJNum > 0 And PNJNum < MAX_NPCS Then
Call PrepareSprite(sprt)
Call AffSurfPic(DD_SpriteSurf(sprt), frmCpnjmouv.skin, 0, DIR_LEFT * (DDSD_Character(sprt).lHeight / 4))
End If
End Sub

Private Sub haut_Click()
Dim sprt As Long
On Error Resume Next
sprt = Npc(PNJNum).sprite
DIRC = DIR_UP
If PNJNum > 0 And PNJNum < MAX_NPCS Then
Call PrepareSprite(sprt)
Call AffSurfPic(DD_SpriteSurf(sprt), frmCpnjmouv.skin, 0, DIR_UP * (DDSD_Character(sprt).lHeight / 4))
End If
End Sub

Private Sub imob_Click()
If mvh.value = Checked Then imob.value = Unchecked: Exit Sub
If imob.value = Unchecked Then
    mvh.Enabled = True
    boucle.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    For i = 1 To 2
        x(i).Enabled = True
        y(i).Enabled = True
    Next i
    defxy1.Enabled = True
    defxy2.Enabled = True
    collco(1).Enabled = True
    collco(2).Enabled = True
    gauche.Enabled = False
    droite.Enabled = False
    haut.Enabled = False
    bas.Enabled = False
    skin.Enabled = False
Else
    mvh.Enabled = False
    boucle.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    For i = 1 To 2
        x(i).Enabled = False
        y(i).Enabled = False
    Next i
    defxy1.Enabled = False
    defxy2.Enabled = False
    collco(1).Enabled = False
    collco(2).Enabled = False
    gauche.Enabled = True
    droite.Enabled = True
    haut.Enabled = True
    bas.Enabled = True
    skin.Enabled = True
End If
End Sub

Private Sub mvh_Click()
If imob.value = Checked Then mvh.value = Checked: Exit Sub
If mvh.value = Unchecked Then
    th.value = Unchecked
    boucle.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    For i = 1 To 2
        x(i).Enabled = True
        y(i).Enabled = True
    Next i
    defxy1.Enabled = True
    defxy2.Enabled = True
    collco(1).Enabled = True
    collco(2).Enabled = True
Else
    If psh.value = Checked Then th.value = Checked
    boucle.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    For i = 1 To 2
        x(i).Enabled = False
        y(i).Enabled = False
    Next i
    defxy1.Enabled = False
    defxy2.Enabled = False
    collco(1).Enabled = False
    collco(2).Enabled = False
End If
End Sub

Private Sub OK_Click()
For i = 0 To 2
If Val(x(i).Text) > 30 Then Call MsgBox("Veuillez mettre des coordonnées en X en chiffre et inférieur à 30"): Exit Sub
If Val(y(i).Text) > 30 Then Call MsgBox("Veuillez mettre des coordonnées en Y en chiffre et inférieur à 30"): Exit Sub
Next i
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x = Val(x(0).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y = Val(y(0).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x1 = Val(x(1).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y1 = Val(y(1).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).x2 = Val(x(2).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).y2 = Val(y(2).Text)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Hasardm = Val(mvh.value)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Hasardp = Val(psh.value)
Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).boucle = Val(boucle.value)
If imob.value = Checked Then Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Imobile = 1 + DIRC Else Map(Player(MyIndex).Map).Npcs(EditorMouvIndex).Imobile = 0
InMouvEditor = False
Unload Me
frmMapProperties.SetFocus
save = 1
Call WriteINI("modif", "carte" & Player(MyIndex).Map, "1", App.Path & "\config.ini")
End Sub

Private Sub psh_Click()
If psh.value = Unchecked Then
    th.value = Unchecked
    Label1.Enabled = True
    Label2.Enabled = True
    x(0).Enabled = True
    y(0).Enabled = True
    defxy.Enabled = True
    collco(0).Enabled = True
Else
    If mvh.value = Checked Then th.value = Checked
    Label1.Enabled = False
    Label2.Enabled = False
    x(0).Enabled = False
    y(0).Enabled = False
    defxy.Enabled = False
    collco(0).Enabled = False
End If
End Sub

Private Sub scrlSkinX_Change()
    skin.Left = scrlSkinX.value
End Sub

Private Sub scrlSkinY_Change()
    skin.Top = scrlSkinY.value
End Sub

Private Sub th_Click()
If th.value = Unchecked Then
    psh.value = Unchecked
    mvh.value = Unchecked
    boucle.Enabled = True
    Label1.Enabled = True
    Label2.Enabled = True
    For i = 0 To 2
        x(i).Enabled = True
        y(i).Enabled = True
    Next i
    defxy.Enabled = True
    Label3.Enabled = True
    Label4.Enabled = True
    Label5.Enabled = True
    Label6.Enabled = True
    defxy1.Enabled = True
    defxy2.Enabled = True
    collco(0).Enabled = True
    collco(1).Enabled = True
    collco(2).Enabled = True
Else
    psh.value = Checked
    mvh.value = Checked
    boucle.Enabled = False
    Label1.Enabled = False
    Label2.Enabled = False
    For i = 0 To 2
        x(i).Enabled = False
        y(i).Enabled = False
    Next i
    defxy.Enabled = False
    Label3.Enabled = False
    Label4.Enabled = False
    Label5.Enabled = False
    Label6.Enabled = False
    defxy1.Enabled = False
    defxy2.Enabled = False
    collco(0).Enabled = False
    collco(1).Enabled = False
    collco(2).Enabled = False
End If
End Sub
