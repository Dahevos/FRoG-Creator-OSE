VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmoptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6255
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Enregistrer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   24
      ToolTipText     =   "Quitte la fenêtre d'édition et enregistre le PNJ"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer le PNJ"
      Top             =   5640
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      TabMaxWidth     =   2778
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Options Générales"
      TabPicture(0)   =   "frmoptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLines"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkAutoScroll"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlBltText"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chksound"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkmusic"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chknpcdamage"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkplayerdamage"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chknpcbar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkbubblebar"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chknpcname"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkplayername"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkplayerbar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chknobj"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkLowEffect"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Config. du Jeu"
      TabPicture(1)   =   "frmoptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label11"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ps"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "script"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "defl"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "pm"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "pv"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "site"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "nom"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Config. du Serveur"
      TabPicture(2)   =   "frmoptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label34"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label30"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label28"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label27"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label26"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label24"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label23"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label21"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label20"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label19"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label18"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label29"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label1"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "mq"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "ms"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "mo"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "mjg"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "me"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "mn"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "mj"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "mpnj"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "mm"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "moc"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "mg"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "mc"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "motd"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "Config. des Classes"
      TabPicture(3)   =   "frmoptions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label33"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "editcls"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "clase"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "nbcls"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      Begin VB.CheckBox chkLowEffect 
         Caption         =   "Désactiver les effets avancés"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   66
         Top             =   4800
         Width           =   2325
      End
      Begin VB.CheckBox chknobj 
         Caption         =   "Nom des objets aux sol (quand la souris les survole)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   65
         ToolTipText     =   "Petit barre afficher au dessus de vous"
         Top             =   1560
         Value           =   1  'Checked
         Width           =   3480
      End
      Begin VB.CheckBox chkplayerbar 
         Caption         =   "Mini barre de vie"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   59
         ToolTipText     =   "Petit barre afficher au dessu de vous"
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkplayername 
         Caption         =   "Nom"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   58
         Top             =   840
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chknpcname 
         Caption         =   "Noms"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   57
         Top             =   2040
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkbubblebar 
         Caption         =   "Bulles de dialogue"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   56
         Top             =   3720
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chknpcbar 
         Caption         =   "Afficher leur mini barre de vie"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   55
         ToolTipText     =   "Petit barre afficher au dessu des PNJ"
         Top             =   2520
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerdamage 
         Caption         =   "Dégâts affichés au dessus de la tête"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   54
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin VB.CheckBox chknpcdamage 
         Caption         =   "Dégâts affichés au dessus de la tête"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   53
         Top             =   2280
         Value           =   1  'Checked
         Width           =   2685
      End
      Begin VB.CheckBox chkmusic 
         Caption         =   "Musique"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   52
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chksound 
         Caption         =   "Effets sonores"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1320
         TabIndex        =   51
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   1320
         Max             =   20
         Min             =   4
         TabIndex        =   50
         Top             =   4200
         Value           =   6
         Width           =   2295
      End
      Begin VB.CheckBox chkAutoScroll 
         Caption         =   "Défilement automatique"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   49
         Top             =   4560
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.TextBox nbcls 
         Height          =   285
         Left            =   -72360
         TabIndex        =   21
         Text            =   "3"
         ToolTipText     =   "Défaut = 3"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox clase 
         Height          =   315
         ItemData        =   "frmoptions.frx":0070
         Left            =   -72960
         List            =   "frmoptions.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton editcls 
         Caption         =   "Editer la classe sélectionée"
         Height          =   300
         Left            =   -73080
         TabIndex        =   23
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox motd 
         Height          =   1575
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "frmoptions.frx":0074
         Top             =   3720
         Width           =   6015
      End
      Begin VB.TextBox mc 
         Height          =   285
         Left            =   -73320
         TabIndex        =   13
         Text            =   "255"
         ToolTipText     =   "Défaut = 255"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox mg 
         Height          =   285
         Left            =   -70200
         TabIndex        =   15
         Text            =   "20"
         ToolTipText     =   "Défaut = 20"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox moc 
         Height          =   285
         Left            =   -70200
         TabIndex        =   14
         Text            =   "20"
         ToolTipText     =   "Défaut = 20"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox mm 
         Height          =   285
         Left            =   -73320
         TabIndex        =   11
         Text            =   "1000"
         ToolTipText     =   "Défaut = 1000"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox mpnj 
         Height          =   285
         Left            =   -73320
         TabIndex        =   10
         Text            =   "1000"
         ToolTipText     =   "Défaut = 1000"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox mj 
         Height          =   285
         Left            =   -73320
         TabIndex        =   8
         Text            =   "50"
         ToolTipText     =   "Défaut = 50"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox mn 
         Height          =   285
         Left            =   -70200
         TabIndex        =   18
         Text            =   "100"
         ToolTipText     =   "Défaut = 500"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox me 
         Height          =   285
         Left            =   -70200
         TabIndex        =   17
         Text            =   "10"
         ToolTipText     =   "Défaut = 10"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox mjg 
         Height          =   285
         Left            =   -70200
         TabIndex        =   16
         Text            =   "20"
         ToolTipText     =   "Défaut = 20"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox mo 
         Height          =   285
         Left            =   -73320
         TabIndex        =   9
         Text            =   "1000"
         ToolTipText     =   "Défaut = 1000"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox ms 
         Height          =   285
         Left            =   -73320
         TabIndex        =   12
         Text            =   "1000"
         ToolTipText     =   "Défaut = 1000"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox mq 
         Height          =   285
         Left            =   -70200
         TabIndex        =   19
         Text            =   "100"
         ToolTipText     =   "Défaut = 500"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox nom 
         Height          =   285
         Left            =   -71280
         TabIndex        =   1
         Text            =   "FRoG Creator"
         ToolTipText     =   "ex : Frog Creator"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox site 
         Height          =   285
         Left            =   -71280
         TabIndex        =   2
         Text            =   "www.frogcreator.fr"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox pv 
         Height          =   285
         Left            =   -71280
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Vitesse de régénération des points de vie"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox pm 
         Height          =   285
         Left            =   -71280
         TabIndex        =   4
         Text            =   "1"
         ToolTipText     =   "Vitesse de régénération des points de magie"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox defl 
         Height          =   285
         Left            =   -71280
         TabIndex        =   6
         Text            =   "1"
         ToolTipText     =   "1 = oui  0 = non"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox script 
         Height          =   285
         Left            =   -71280
         TabIndex        =   7
         Text            =   "1"
         ToolTipText     =   "1 = oui  0 = non"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox ps 
         Height          =   285
         Left            =   -71280
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Vitesse de régénération des points spéciale"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de ligne écrite sur l'écran: 6"
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
         Left            =   1320
         TabIndex        =   64
         Top             =   4020
         Width           =   2220
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage du Joueur-"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   63
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage des PNJ-"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   62
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Musique/Sons-"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   61
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage du Chat-"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1680
         TabIndex        =   60
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   $"frmoptions.frx":00FE
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -74640
         TabIndex        =   48
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   $"frmoptions.frx":01B3
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -74640
         TabIndex        =   47
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label33 
         Caption         =   "Maximum de classes (de 0 à .....)  :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label29 
         Caption         =   "Message envoyé aux joueurs quand ils se connecteront sur le jeu :"
         Height          =   375
         Left            =   -74880
         TabIndex        =   45
         Top             =   3360
         Width           =   6015
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Max de Cartes :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   44
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Max de Guildes :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   43
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Max d'Objets/Carte :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   42
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Max de Magasins :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   41
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Max de PNJ :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   40
         Top             =   1920
         Width           =   960
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Max d'Objets :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   39
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Max de joueurs :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   38
         Top             =   1200
         Width           =   1170
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Max joueurs/Guilde :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   37
         Top             =   1920
         Width           =   1470
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Max Emoticones :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   36
         Top             =   2280
         Width           =   1260
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Max Niveaux :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   35
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Max de Sorts :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   34
         Top             =   2640
         Width           =   1020
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Max Quêtes :"
         Height          =   195
         Left            =   -71880
         TabIndex        =   33
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label10 
         Caption         =   "Nom de votre jeu :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Site de votre jeu (facultatif) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   31
         Top             =   1560
         Width           =   1980
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Pv régénérés (toute les 3 secondes environ) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   30
         Top             =   1920
         Width           =   3210
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Pm régénérés (toute les 3 secondes environ) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   29
         Top             =   2280
         Width           =   3240
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Défilement des cartes (1 = oui, 0 = non) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   3000
         Width           =   2865
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Activer les scripts (1 = oui, 0 = non) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   27
         Top             =   3360
         Width           =   2565
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ps régénérés (toute les 3 secondes environ) :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   26
         Top             =   2640
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmoptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkLowEffect_Click()
    WriteINI "CONFIG", "LowEffect", chkLowEffect.value, App.Path & "\Config\Account.ini"
    AccOpt.LowEffect = CBool(chkLowEffect.value)
End Sub

Private Sub chknobj_Click()
    WriteINI "CONFIG", "NomObjet", chknobj.value, App.Path & "\Config\Account.ini"
    AccOpt.NomObjet = CBool(chknobj.value)
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.value, App.Path & "\Config\Account.ini"
    AccOpt.Sound = CBool(chksound.value)
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.value, App.Path & "\Config\Account.ini"
    AccOpt.SpeechBubbles = CBool(chkbubblebar.value)
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.value, App.Path & "\Config\Account.ini"
    AccOpt.NpcBar = CBool(chknpcbar.value)
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.value, App.Path & "\Config\Account.ini"
    AccOpt.NpcDamage = CBool(chknpcdamage.value)
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.value, App.Path & "\Config\Account.ini"
    AccOpt.NpcName = CBool(chknpcname.value)
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.value, App.Path & "\Config\Account.ini"
    AccOpt.PlayBar = CBool(chkplayerbar.value)
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.value, App.Path & "\Config\Account.ini"
    AccOpt.PlayDamage = CBool(chkplayerdamage.value)
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.value, App.Path & "\Config\Account.ini"
    AccOpt.Autoscroll = CBool(chkAutoScroll.value)
End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    MAX_BLT_LINE = scrlBltText.value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "Nbr de ligne écrite sur l'écran: " & scrlBltText.value
End Sub
Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.value, App.Path & "\Config\Account.ini"
    AccOpt.Music = CBool(chkmusic.value)
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
frmmsg.Show
Call OptionSave
Call InitAccountOpt
Unload Me
End Sub

Private Sub editcls_Click()
    If clase.ListIndex < 0 Or clase.ListIndex > Val(nbcls.Text) Then Exit Sub
    classe = clase.ListIndex
    frmclasseseditor.nom.Text = ReadINI("CLASS", "Name", App.Path & "\Classes\Class" & classe & ".ini")
    frmclasseseditor.scrlhom.value = Val(ReadINI("CLASS", "MaleSprite", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.scrlfem.value = Val(ReadINI("CLASS", "FemaleSprite", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.numsf.Caption = Val(frmclasseseditor.scrlfem.value)
    frmclasseseditor.numsh.Caption = Val(frmclasseseditor.scrlhom.value)
    frmclasseseditor.force.Text = Val(ReadINI("CLASS", "STR", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.def.Text = Val(Val(ReadINI("CLASS", "DEF", App.Path & "\Classes\Class" & classe & ".ini")))
    frmclasseseditor.vit.Text = Val(ReadINI("CLASS", "SPEED", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.magi.Text = Val(ReadINI("CLASS", "MAGI", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.carted.Text = Val(ReadINI("CLASS", "MAP", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.xd.Text = Val(ReadINI("CLASS", "X", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.yd.Text = Val(ReadINI("CLASS", "Y", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.arme.Text = Val(ReadINI("STARTUP", "Weapon", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.bouclier.Text = Val(ReadINI("STARTUP", "Shield", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.armure.Text = Val(ReadINI("STARTUP", "Armor", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.caske.Text = Val(ReadINI("STARTUP", "Helmet", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.ajf.Text = Val(ReadINI("CLASSCHANGE", "AddStr", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.ajd.Text = Val(ReadINI("CLASSCHANGE", "AddDef", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.ajv.Text = Val(ReadINI("CLASSCHANGE", "AddSpeed", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.ajm.Text = Val(ReadINI("CLASSCHANGE", "AddMagi", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.cartem.Text = Val(ReadINI("DEATH", "Map", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.xm.Text = Val(ReadINI("DEATH", "x", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.ym.Text = Val(ReadINI("DEATH", "y", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.Lock.value = Val(ReadINI("CLASS", "Locked", App.Path & "\Classes\Class" & classe & ".ini"))
    frmclasseseditor.homme.Height = 48
    frmclasseseditor.femme.Height = 48
    If frmclasseseditor.homme.Height <= 0 Then frmclasseseditor.homme.Height = 48: frmclasseseditor.femme.Height = 48
    frmclasseseditor.mal.Height = (48 * Screen.TwipsPerPixelY) + 44
    frmclasseseditor.fem.Height = (48 * Screen.TwipsPerPixelY) + 44
On Error Resume Next
    Call PrepareSprite(frmclasseseditor.scrlfem.value)
    Call AffSurfPic(DD_SpriteSurf(frmclasseseditor.scrlfem.value), frmclasseseditor.femme, 0, 0)
    Call AffSurfPic(DD_SpriteSurf(frmclasseseditor.scrlhom.value), frmclasseseditor.homme, 0, 0)
    frmclasseseditor.Show
End Sub

Private Sub Form_Load()
Dim i As Long
For i = 0 To Val(nbcls.Text)
    Call clase.AddItem("Classe" & i, i)
Next i

motd.Text = "Bienvenue dans la version " & App.Major & "." & App.Minor & "." & App.Revision & " de FRoG Creator, si vous rencontrez un problème ou un bug veuillez le rapporter sur frogcreator.fr"
End Sub

Private Sub nbcls_Change()
Dim i As Long
Call clase.Clear
For i = 0 To Val(nbcls.Text)
    Call clase.AddItem("Classe" & i, i)
Next i
End Sub
