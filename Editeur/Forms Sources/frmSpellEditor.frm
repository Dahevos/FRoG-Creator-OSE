VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSpellEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer un sort"
   ClientHeight    =   7560
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12000
   DrawMode        =   14  'Copy Pen
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1800
      Top             =   120
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7395
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13044
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1587
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sort"
      TabPicture(0)   =   "frmSpellEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSound"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblVitalMod"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblRange"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblSpellAnim"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblSpellTime"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSpellDone"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label11"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "info"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "scrlSound"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "scrlVitalMod"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbType"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdCancel"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdOk"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "scrlRange"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "scrlSpellAnim"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "scrlSpellTime"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "scrlSpellDone"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "picSpell"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "HScroll1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "CheckSpell"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "picSpellIco"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "picPic"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "VScroll1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      Begin VB.VScrollBar VScroll1 
         Height          =   2760
         LargeChange     =   10
         Left            =   3120
         Max             =   464
         TabIndex        =   42
         Top             =   4440
         Width           =   255
      End
      Begin VB.PictureBox picPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   240
         ScaleHeight     =   183
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   41
         ToolTipText     =   "Sélectionner une image pour l'objet"
         Top             =   4440
         Width           =   2880
      End
      Begin VB.PictureBox picSpellIco 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3960
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   39
         Top             =   5280
         Width           =   480
      End
      Begin VB.CheckBox CheckSpell 
         Caption         =   "BigSpells"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6000
         TabIndex        =   38
         Top             =   4020
         Width           =   1995
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         LargeChange     =   10
         Left            =   6000
         Max             =   10000
         TabIndex        =   35
         Top             =   1260
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualiser"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   220
         Left            =   9420
         TabIndex        =   33
         Top             =   4020
         Width           =   1095
      End
      Begin VB.PictureBox picSpell 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   10920
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   32
         Top             =   3720
         Width           =   480
      End
      Begin VB.HScrollBar scrlSpellDone 
         Height          =   270
         LargeChange     =   2
         Left            =   6000
         Max             =   10
         Min             =   1
         TabIndex        =   31
         Top             =   5220
         Value           =   1
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSpellTime 
         Height          =   270
         LargeChange     =   10
         Left            =   6000
         Max             =   500
         Min             =   40
         TabIndex        =   30
         Top             =   4620
         Value           =   40
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSpellAnim 
         Height          =   270
         LargeChange     =   10
         Left            =   6000
         Max             =   2000
         TabIndex        =   29
         Top             =   3720
         Value           =   1
         Width           =   4515
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   270
         LargeChange     =   5
         Left            =   6000
         Max             =   30
         Min             =   1
         TabIndex        =   25
         Top             =   3060
         Value           =   1
         Width           =   5655
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8520
         TabIndex        =   22
         ToolTipText     =   "Quitte la fenêtre d'édition et enregistre le sort "
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   10080
         TabIndex        =   21
         ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer le sort "
         Top             =   6840
         Width           =   1455
      End
      Begin VB.ComboBox cmbType 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmSpellEditor.frx":001C
         Left            =   6000
         List            =   "frmSpellEditor.frx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Sélectionner un type de sort"
         Top             =   720
         Width           =   5655
      End
      Begin VB.HScrollBar scrlVitalMod 
         Height          =   270
         LargeChange     =   10
         Left            =   6000
         Max             =   10000
         TabIndex        =   14
         Top             =   1860
         Width           =   5655
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   270
         LargeChange     =   10
         Left            =   6000
         Max             =   100
         TabIndex        =   13
         Top             =   2460
         Width           =   5655
      End
      Begin VB.Frame Frame1 
         Caption         =   "Requis de l'utilisation"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   5175
         Begin VB.HScrollBar scrlLevelReq 
            Height          =   270
            Left            =   120
            Max             =   500
            TabIndex        =   8
            Top             =   600
            Value           =   1
            Width           =   4935
         End
         Begin VB.HScrollBar scrlCost 
            Height          =   270
            Left            =   120
            Max             =   1000
            TabIndex        =   7
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Niveau Requis:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Niveau requis pour utiliser le sort"
            Top             =   360
            Width           =   945
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PM Requis:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Point(s) de magie requis pour utiliser le sort"
            Top             =   960
            Width           =   705
         End
         Begin VB.Label lblLevelReq 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Administrateur Seulement"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1200
            TabIndex        =   10
            ToolTipText     =   "Niveau requis pour utiliser le sort"
            Top             =   360
            Width           =   1605
         End
         Begin VB.Label lblCost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   960
            TabIndex        =   9
            ToolTipText     =   "Point(s) de magie requis pour utiliser le sort"
            Top             =   960
            Width           =   75
         End
      End
      Begin VB.Frame info 
         Caption         =   "Informations"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   5175
         Begin VB.CheckBox chkArea 
            Caption         =   "Affecte les alentours"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Le sort affectera les joueurs ou PNJ qui se trouvent à la distance sélectionner autour de sont lanceur si la case est cocher"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.ComboBox cmbClassReq 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            ItemData        =   "frmSpellEditor.frx":0147
            Left            =   120
            List            =   "frmSpellEditor.frx":0149
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Classe requise pour utiliser le sort"
            Top             =   1080
            Width           =   4905
         End
         Begin VB.TextBox txtName 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Nom du sort"
            Top             =   480
            Width           =   4875
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Classe requise :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Nom :"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label Label11 
         Caption         =   "Icone :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   40
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de points perdu :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   37
         ToolTipText     =   "Nombre de points que va perdre/gagner le joueur dans toutes ses caratéristiques"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8040
         TabIndex        =   36
         ToolTipText     =   "Modifie les points correspondants aux types de cette valeur"
         Top             =   1080
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblSpellDone 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cycle Animation 1 Fois"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6000
         TabIndex        =   28
         ToolTipText     =   "Nombre de fois que l'animation va se répéter"
         Top             =   4980
         Width           =   1455
      End
      Begin VB.Label lblSpellTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Temps: 40"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6000
         TabIndex        =   27
         ToolTipText     =   "Intervalle de l'animation plus le chiffre et grand plus l'animation est lente"
         Top             =   4380
         Width           =   660
      End
      Begin VB.Label lblSpellAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: 0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6000
         TabIndex        =   26
         ToolTipText     =   "Numéros de l'animation du sort"
         Top             =   3480
         Width           =   2115
      End
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6720
         TabIndex        =   24
         ToolTipText     =   "Distance en cases de l'effet du sort "
         Top             =   2820
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Distance:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   23
         ToolTipText     =   "Distance en cases de l'effet du sort "
         Top             =   2820
         Width           =   705
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Type de Sort :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   480
         Width           =   840
      End
      Begin VB.Label lblVitalMod 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8040
         TabIndex        =   19
         ToolTipText     =   "Modifie les points correspondants aux types de cette valeur"
         Top             =   1620
         Width           =   75
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modifie les points de :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   18
         ToolTipText     =   "Modifie les points correspondants aux types de cette valeur"
         Top             =   1620
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Effet Sonore:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   17
         ToolTipText     =   "Nom de l'effet sonore sélectionner"
         Top             =   2220
         Width           =   900
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aucun"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6960
         TabIndex        =   16
         ToolTipText     =   "Nom de l'effet sonore sélectionner"
         Top             =   2220
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmSpellEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Done As Long
Private Time As Long
Private SpellVar As Long

Private Sub CheckSpell_Click()
    If CheckSpell.value = Checked Then
        scrlSpellAnim.Max = MAX_DX_BIGSPELLS
        picSpell.Width = 960
        picSpell.Height = 960
        picSpell.Left = 10680
        picSpell.Top = 3540
        scrlSpellAnim.value = 0
    Else
        scrlSpellAnim.Max = MAX_DX_SPELLS
        picSpell.Width = 480
        picSpell.Height = 480
        picSpell.Left = 10920
        picSpell.Top = 3720
        scrlSpellAnim.value = 0
    End If
    Done = 0
End Sub

Private Sub cmbType_Click()
HScroll1.Visible = False
Label10.Visible = False
Label9.Visible = False

If cmbType.ListIndex = SPELL_TYPE_SCRIPT Then
    Label4.Caption = "Numéro de la case :"
    Label4.ToolTipText = "Numéro de la case de script qui va s'exécuter"
ElseIf cmbType.ListIndex = SPELL_TYPE_PARALY Then
    Label4.Caption = "Temps de paralysie(seconde) :"
    Label4.ToolTipText = "Nombre de seconde(s) que va durée la paralysie"
ElseIf cmbType.ListIndex = SPELL_TYPE_DEFENC Then
    Label4.Caption = "Temps de protéction(seconde) :"
    Label4.ToolTipText = "Nombre de seconde(s) que va durée le bouclier contre les sorts"
ElseIf cmbType.ListIndex = SPELL_TYPE_AMELIO Then
    Label4.Caption = "Durée du gain(seconde) :"
    Label4.ToolTipText = "Nombre de seconde(s) que va durée le gain de points de caractéristique"
    Label10.Caption = "Nombre de points gagner :"
    HScroll1.Visible = True
    Label10.Visible = True
    Label9.Visible = True
ElseIf cmbType.ListIndex = SPELL_TYPE_DECONC Then
    Label4.Caption = "Durée de la perte(seconde) :"
    Label4.ToolTipText = "Nombre de seconde(s) que va durée la perte de points de caractéristique"
    Label10.Caption = "Nombre de points perdu :"
    HScroll1.Visible = True
    Label10.Visible = True
    Label9.Visible = True
Else
    Label4.Caption = "Modifie les points de :"
    Label4.ToolTipText = "Modifie les points correspondants aux types de cette valeur"
End If

End Sub

Private Sub Command1_Click()
    Done = 0
End Sub

Private Sub Form_Load()
    picSpell.Width = 480
    picSpell.Height = 480
    picSpell.Left = 10920
    picSpell.Top = 3720
    Call AffSurfPic(DD_ItemSurf, picSpellIco, EditorItemX * PIC_X, EditorItemY * PIC_Y)
    Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SpellEditorCancel
End Sub

Private Sub HScroll1_Change()
    Label9.Caption = CStr(HScroll1.value)
End Sub

Private Sub picPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorItemX = (x \ PIC_X): EditorItemY = (y \ PIC_Y) + VScroll1.value
    Call AffSurfPic(DD_ItemSurf, picSpellIco, EditorItemX * PIC_X, EditorItemY * PIC_Y)
End Sub

Private Sub picPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorItemX = (x \ PIC_X): EditorItemY = (y \ PIC_Y) + VScroll1.value
    Call AffSurfPic(DD_ItemSurf, picSpellIco, EditorItemX * PIC_X, EditorItemY * PIC_Y)
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = CStr(scrlCost.value)
End Sub

Private Sub scrlLevelReq_Change()
    If CStr(scrlLevelReq.value) = 0 Then lblLevelReq.Caption = "Administrateur Seulement" Else lblLevelReq.Caption = CStr(scrlLevelReq.value)
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.value)
End Sub

Private Sub scrlSound_Change()
If CStr(scrlSound.value) = 0 Then lblSound.Caption = "Aucuns" Else lblSound.Caption = CStr(scrlSound.value): Call PlaySound("magic" & scrlSound.value & ".wav")
End Sub

Private Sub scrlSpellAnim_Change()
    lblSpellAnim.Caption = "Anim: " & scrlSpellAnim.value
    Done = 0
End Sub

Private Sub scrlSpellDone_Change()
Dim String2 As String
    String2 = "Fois"
    lblSpellDone.Caption = "Cycle Animation " & scrlSpellDone.value & " " & String2
    Done = 0
End Sub

Private Sub scrlSpellTime_Change()
    lblSpellTime.Caption = "Temps: " & scrlSpellTime.value
    Done = 0
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.value)
End Sub

Private Sub cmdOk_Click()
    Call SpellEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call SpellEditorCancel
End Sub

Private Sub Timer1_Timer()
Dim sRECT As RECT
Dim dRECT As RECT
Dim SpellDone As Long
Dim SpellAnim As Long
Dim SpellTime As Long

SpellDone = scrlSpellDone.value
SpellAnim = scrlSpellAnim.value
SpellTime = scrlSpellTime.value

'If SpellAnim <= 0 Then Exit Sub
If Done = SpellDone Then Exit Sub

    With dRECT
        .Top = 0
        .Bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With

    If SpellVar > 10 Then Done = Done + 1: SpellVar = 0
    If GetTickCount > Time + SpellTime Then Time = GetTickCount: SpellVar = SpellVar + 1
    If CheckSpell.value = Checked Then
        Call PrepareBigSpell(SpellAnim)
        If DD_BigSpellAnim(SpellAnim) Is Nothing Then
        Else
            With dRECT
                .Top = 0
                .Bottom = PIC_Y * 2
                .Left = 0
                .Right = PIC_X * 2
            End With
            With sRECT
                .Top = 0 * (PIC_Y * 2)
                .Bottom = .Top + (PIC_Y * 2)
                .Left = SpellVar * (PIC_X * 2)
                .Right = .Left + (PIC_X * 2)
            End With
            Call PrepareBigSpell(SpellAnim)
            Call DD_BigSpellAnim(SpellAnim).BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    Else
        Call PrepareSpell(SpellAnim)
        If DD_SpellAnim(SpellAnim) Is Nothing Then
        Else
            With dRECT
                .Top = 0
                .Bottom = PIC_Y
                .Left = 0
                .Right = PIC_X
            End With
            With sRECT
                .Top = 0 * PIC_Y
                .Bottom = .Top + PIC_Y
                .Left = SpellVar * PIC_X
                .Right = .Left + PIC_X
            End With
            Call PrepareSpell(SpellAnim)
            Call DD_SpellAnim(SpellAnim).BltToDC(picSpell.hDC, sRECT, dRECT)
            picSpell.Refresh
        End If
    End If
End Sub

Private Sub VScroll1_Change()
    Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
End Sub

Private Sub VScroll1_Scroll()
    Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
End Sub
