VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMirage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10020
   ClientLeft      =   1275
   ClientTop       =   1140
   ClientWidth     =   12000
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmMirage.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   660
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   Visible         =   0   'False
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   9000
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   143
      Top             =   3240
      Visible         =   0   'False
      Width           =   2625
      Begin VB.TextBox txtTempsBulles 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   251
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
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
         TabIndex        =   144
         Top             =   5400
         Width           =   2415
      End
      Begin VB.CommandButton CmdoptTouche 
         Caption         =   "Configurer les touches"
         Height          =   255
         Left            =   120
         TabIndex        =   145
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CheckBox chknobj 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nom des objets aux sol (quand la souris le survole)"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   159
         ToolTipText     =   "Petite barre affichée au dessus de vous"
         Top             =   960
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerbar 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   158
         Top             =   720
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkplayername 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   157
         Top             =   240
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chknpcname 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   156
         Top             =   1440
         Value           =   1  'Checked
         Width           =   765
      End
      Begin VB.CheckBox chkbubblebar 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   155
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chknpcbar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Affichés leur mini barre de vie"
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
         Left            =   120
         TabIndex        =   154
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.CheckBox chkplayerdamage 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   153
         Top             =   480
         Value           =   1  'Checked
         Width           =   2565
      End
      Begin VB.CheckBox chknpcdamage 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   152
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkmusic 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   151
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chksound 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   150
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox chkAutoScroll 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   149
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1845
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualiser le thème"
         Height          =   255
         Left            =   120
         TabIndex        =   148
         Top             =   4920
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.HScrollBar scrlBltText 
         Height          =   255
         Left            =   240
         Max             =   20
         Min             =   4
         TabIndex        =   147
         Top             =   4125
         Value           =   6
         Width           =   2055
      End
      Begin VB.CheckBox chkLowEffect 
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   146
         Top             =   4680
         Width           =   2325
      End
      Begin VB.Label lblBulle 
         BackStyle       =   0  'Transparent
         Caption         =   "Temps d'affichage des bulles:"
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
         TabIndex        =   250
         Top             =   3360
         Width           =   2175
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
         Left            =   120
         TabIndex        =   164
         Top             =   3960
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
         Left            =   -120
         TabIndex        =   163
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label14 
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
         Left            =   -120
         TabIndex        =   162
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label18 
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
         Left            =   0
         TabIndex        =   161
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Affichage des NPCs-"
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
         Left            =   0
         TabIndex        =   160
         Top             =   1275
         Width           =   2655
      End
   End
   Begin VB.PictureBox pictMetier 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   3960
      Picture         =   "frmMirage.frx":74F2
      ScaleHeight     =   1965
      ScaleWidth      =   3600
      TabIndex        =   241
      Top             =   600
      Visible         =   0   'False
      Width           =   3630
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   249
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label lblOublierMetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Oublier le Metier"
         Height          =   255
         Left            =   1440
         TabIndex        =   248
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblendmetier 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3390
         TabIndex        =   247
         Top             =   1935
         Width           =   375
      End
      Begin VB.Label lblmetierEnd 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fermer"
         Height          =   255
         Left            =   2760
         TabIndex        =   245
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   244
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   243
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label lblmetier 
         BackStyle       =   0  'Transparent
         Caption         =   "Label41"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   242
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.PictureBox PicMenuQuitter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   3480
      Picture         =   "frmMirage.frx":1E5A4
      ScaleHeight     =   1965
      ScaleWidth      =   3600
      TabIndex        =   212
      Top             =   2640
      Visible         =   0   'False
      Width           =   3600
      Begin VB.Label lblCdP 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   216
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDeco 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   215
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblQuitter 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   0
         TabIndex        =   214
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3360
         TabIndex        =   213
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox pictTouche 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4035
      Left            =   3120
      ScaleHeight     =   267
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   165
      Top             =   960
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmdOTA 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   6840
         TabIndex        =   185
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox cbth 
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
         Left            =   930
         TabIndex        =   184
         Text            =   "Combo1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cbtb 
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
         Left            =   930
         TabIndex        =   183
         Text            =   "Combo1"
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cbtg 
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
         Left            =   930
         TabIndex        =   182
         Text            =   "Combo1"
         Top             =   720
         Width           =   2775
      End
      Begin VB.ComboBox cbtd 
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
         Left            =   930
         TabIndex        =   181
         Text            =   "Combo1"
         Top             =   960
         Width           =   2775
      End
      Begin VB.ComboBox cbta 
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
         Left            =   930
         TabIndex        =   180
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox cbtc 
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
         Left            =   930
         TabIndex        =   209
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton cmdOTO 
         Caption         =   "Ok"
         Height          =   255
         Left            =   6120
         TabIndex        =   186
         Top             =   3720
         Width           =   735
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   0
         Left            =   4905
         TabIndex        =   179
         Text            =   "Combo1"
         Top             =   225
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   1
         Left            =   4905
         TabIndex        =   178
         Text            =   "Combo1"
         Top             =   480
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   2
         Left            =   4905
         TabIndex        =   177
         Text            =   "Combo1"
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   3
         Left            =   4905
         TabIndex        =   176
         Text            =   "Combo1"
         Top             =   960
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   4
         Left            =   4905
         TabIndex        =   175
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   5
         Left            =   4905
         TabIndex        =   174
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   6
         Left            =   4905
         TabIndex        =   173
         Text            =   "Combo1"
         Top             =   1695
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   7
         Left            =   4905
         TabIndex        =   172
         Text            =   "Combo1"
         Top             =   1935
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   8
         Left            =   4905
         TabIndex        =   171
         Text            =   "Combo1"
         Top             =   2175
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   9
         Left            =   4905
         TabIndex        =   170
         Text            =   "Combo1"
         Top             =   2415
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   10
         Left            =   4905
         TabIndex        =   169
         Text            =   "Combo1"
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   11
         Left            =   4905
         TabIndex        =   168
         Text            =   "Combo1"
         Top             =   2895
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   12
         Left            =   4905
         TabIndex        =   167
         Text            =   "Combo1"
         Top             =   3135
         Width           =   2655
      End
      Begin VB.ComboBox cbtr 
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
         Index           =   13
         Left            =   4905
         TabIndex        =   166
         Text            =   "Combo1"
         Top             =   3375
         Width           =   2655
      End
      Begin VB.ComboBox cbtra 
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
         Left            =   930
         TabIndex        =   211
         Text            =   "Combo1"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.ComboBox cbtac 
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
         Left            =   930
         TabIndex        =   218
         Text            =   "Combo1"
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "Action :"
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
         TabIndex        =   217
         Top             =   1950
         Width           =   735
      End
      Begin VB.Label Label39 
         BackStyle       =   0  'Transparent
         Caption         =   "Ramasser :"
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
         TabIndex        =   210
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Courir :"
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
         TabIndex        =   208
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "- Touche de Jeu -"
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
         Left            =   0
         TabIndex        =   207
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "- Touche de Racourci -"
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
         Left            =   3960
         TabIndex        =   206
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Haut :"
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
         TabIndex        =   205
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Bas :"
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
         TabIndex        =   204
         Top             =   510
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Gauche :"
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
         TabIndex        =   203
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Droite :"
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
         TabIndex        =   202
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Attaque :"
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
         TabIndex        =   201
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 1 :"
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
         Left            =   3960
         TabIndex        =   200
         Top             =   270
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 2 :"
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
         Left            =   3960
         TabIndex        =   199
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 3 :"
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
         Left            =   3960
         TabIndex        =   198
         Top             =   750
         Width           =   855
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 4 :"
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
         Left            =   3960
         TabIndex        =   197
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 8 :"
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
         Left            =   3960
         TabIndex        =   196
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 7 :"
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
         Left            =   3960
         TabIndex        =   195
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 6 :"
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
         Left            =   3960
         TabIndex        =   194
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 5 :"
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
         Left            =   3960
         TabIndex        =   193
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 12 :"
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
         Left            =   3960
         TabIndex        =   192
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 11 :"
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
         Left            =   3960
         TabIndex        =   191
         Top             =   2670
         Width           =   855
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 10 :"
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
         Left            =   3960
         TabIndex        =   190
         Top             =   2430
         Width           =   855
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 9 :"
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
         Left            =   3960
         TabIndex        =   189
         Top             =   2190
         Width           =   855
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 14 :"
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
         Left            =   3960
         TabIndex        =   188
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Raccourci 13 :"
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
         Left            =   3960
         TabIndex        =   187
         Top             =   3150
         Width           =   855
      End
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   7125
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   141
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   6585
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   140
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   6045
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   139
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   5505
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   138
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   4965
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   137
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   4425
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   136
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   3885
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   135
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   3345
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   134
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   2805
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   133
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   2265
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   132
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   1725
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   131
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   1185
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   130
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   645
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   129
      Top             =   9315
      Width           =   480
   End
   Begin VB.PictureBox picRac 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   105
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   128
      Top             =   9315
      Width           =   480
   End
   Begin VB.ComboBox Canal 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMirage.frx":21156
      Left            =   120
      List            =   "frmMirage.frx":21166
      Locked          =   -1  'True
      TabIndex        =   120
      Text            =   "Carte"
      Top             =   8760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtMyTextBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1335
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   111
      Top             =   8760
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.PictureBox itmDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   9360
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   98
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Label descName 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   110
         ToolTipText     =   "Nom de l'objet"
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Requière-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   109
         ToolTipText     =   "Force/défense/vitesse requise pour équipper l'objet"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label descStr 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Strength"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   108
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label descDef 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Defence"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   107
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label descSpeed 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   360
         TabIndex        =   106
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Donne-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   105
         ToolTipText     =   "Se que vous apporte l'objet"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label descHpMp 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "HP: XXXX MP: XXXX SP: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   104
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label descSD 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Str: XXXX Def: XXXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   103
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label desc 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   120
         TabIndex        =   102
         ToolTipText     =   "Description de l'objet"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "-Description-"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   360
         TabIndex        =   101
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label descMS 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Magi: XXXXX Speed: XXXX"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   0
         TabIndex        =   100
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Usure 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Usure : XXXX/XXXX"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   99
         Top             =   2160
         Width           =   2655
      End
   End
   Begin VB.Timer quetetimersec 
      Enabled         =   0   'False
      Left            =   9240
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   7800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picturesprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   68
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7380
      Left            =   0
      ScaleHeight     =   492
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11040
      Begin VB.Timer sync 
         Interval        =   250
         Left            =   6720
         Top             =   0
      End
      Begin VB.Frame fra_fenetre 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   2985
         Left            =   8400
         TabIndex        =   2
         Top             =   4320
         Width           =   2595
         Begin VB.PictureBox tmpsquete 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   465
            Left            =   720
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   81
            TabIndex        =   62
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
            Begin VB.Label seconde 
               BackStyle       =   0  'Transparent
               Caption         =   "00"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   600
               TabIndex        =   64
               ToolTipText     =   "Secondes restante avant la fin de la quête en cour"
               Top             =   0
               Width           =   450
            End
            Begin VB.Label minute 
               BackStyle       =   0  'Transparent
               Caption         =   "00:"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   63
               ToolTipText     =   "Minutes restante avant la fin de la quête en cour"
               Top             =   0
               Width           =   600
            End
         End
         Begin VB.PictureBox picInv3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   120
            ScaleHeight     =   169
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   161
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
            Begin VB.PictureBox Up 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   720
               Picture         =   "frmMirage.frx":21186
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   51
               Top             =   2235
               Width           =   270
            End
            Begin VB.PictureBox Down 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1080
               Picture         =   "frmMirage.frx":2141E
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   50
               Top             =   2235
               Width           =   270
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   330
               Left            =   2640
               Max             =   100
               TabIndex        =   49
               Top             =   2400
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.PictureBox Picture8 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   1935
               Left            =   120
               ScaleHeight     =   129
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   137
               TabIndex        =   48
               Top             =   120
               Width           =   2055
               Begin VB.PictureBox Picture9 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   7935
                  Left            =   120
                  ScaleHeight     =   529
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   126
                  TabIndex        =   59
                  Top             =   0
                  Width           =   1890
                  Begin VB.PictureBox picInv 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Segoe UI"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   480
                     Index           =   0
                     Left            =   120
                     ScaleHeight     =   32
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   32
                     TabIndex        =   60
                     Top             =   120
                     Width           =   480
                  End
                  Begin VB.Shape IDAD 
                     BorderColor     =   &H00008000&
                     BorderWidth     =   3
                     Height          =   510
                     Left            =   0
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   510
                  End
                  Begin VB.Shape EquipS 
                     BorderColor     =   &H0000FFFF&
                     BorderWidth     =   3
                     Height          =   540
                     Index           =   4
                     Left            =   0
                     Top             =   0
                     Width           =   540
                  End
                  Begin VB.Shape EquipS 
                     BorderColor     =   &H0000FFFF&
                     BorderWidth     =   3
                     Height          =   540
                     Index           =   0
                     Left            =   0
                     Top             =   0
                     Width           =   540
                  End
                  Begin VB.Shape EquipS 
                     BorderColor     =   &H0000FFFF&
                     BorderWidth     =   3
                     Height          =   540
                     Index           =   1
                     Left            =   -360
                     Top             =   120
                     Width           =   540
                  End
                  Begin VB.Shape EquipS 
                     BorderColor     =   &H0000FFFF&
                     BorderWidth     =   3
                     Height          =   540
                     Index           =   2
                     Left            =   0
                     Top             =   120
                     Width           =   540
                  End
                  Begin VB.Shape EquipS 
                     BorderColor     =   &H0000FFFF&
                     BorderWidth     =   3
                     Height          =   540
                     Index           =   3
                     Left            =   0
                     Top             =   0
                     Width           =   540
                  End
                  Begin VB.Shape SelectedItem 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   2
                     Height          =   525
                     Left            =   105
                     Top             =   105
                     Width           =   525
                  End
               End
            End
            Begin VB.Line Line2 
               X1              =   4
               X2              =   171
               Y1              =   144
               Y2              =   144
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   160
               Y1              =   144
               Y2              =   144
            End
            Begin VB.Label lblDropItem 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Jeter"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   1320
               TabIndex        =   53
               Top             =   2265
               Width           =   555
            End
            Begin VB.Label lblUseItem 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Utiliser"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   15
               TabIndex        =   52
               Top             =   2265
               Width           =   690
            End
         End
         Begin VB.PictureBox picPlayerSpells 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   120
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   161
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
            Begin VB.PictureBox Picture18 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1200
               Picture         =   "frmMirage.frx":216A9
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   127
               Top             =   2250
               Width           =   270
            End
            Begin VB.PictureBox Picture17 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   840
               Picture         =   "frmMirage.frx":21934
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   126
               Top             =   2250
               Width           =   270
            End
            Begin VB.PictureBox Picture13 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   2175
               Left            =   0
               ScaleHeight     =   2175
               ScaleWidth      =   2415
               TabIndex        =   123
               Top             =   0
               Width           =   2415
               Begin VB.PictureBox Picture11 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   7935
                  Left            =   240
                  ScaleHeight     =   529
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   126
                  TabIndex        =   124
                  Top             =   0
                  Width           =   1890
                  Begin VB.PictureBox picspell 
                     Appearance      =   0  'Flat
                     AutoRedraw      =   -1  'True
                     BackColor       =   &H00000000&
                     BorderStyle     =   0  'None
                     BeginProperty Font 
                        Name            =   "Segoe UI"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   480
                     Index           =   0
                     Left            =   120
                     ScaleHeight     =   32
                     ScaleMode       =   3  'Pixel
                     ScaleWidth      =   32
                     TabIndex        =   125
                     Top             =   120
                     Width           =   480
                  End
                  Begin VB.Shape SDAD 
                     BorderColor     =   &H00008000&
                     BorderWidth     =   3
                     Height          =   510
                     Left            =   105
                     Top             =   105
                     Visible         =   0   'False
                     Width           =   510
                  End
               End
            End
         End
         Begin VB.PictureBox picWhosOnline 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   2385
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
            Begin VB.ListBox lstOnline 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2190
               ItemData        =   "frmMirage.frx":21BCC
               Left            =   75
               List            =   "frmMirage.frx":21BCE
               TabIndex        =   45
               Top             =   75
               Width           =   2220
            End
         End
         Begin VB.Frame fraCarte 
            BackColor       =   &H80000009&
            BorderStyle     =   0  'None
            Height          =   2415
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
            Begin VB.Image imgCarte 
               Height          =   2295
               Left            =   420
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.PictureBox picEquip 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   120
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   159
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
            Begin VB.PictureBox AmuletImage2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   1680
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   27
               Top             =   120
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox AmuletImage 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   28
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture14 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   480
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   25
               Top             =   1920
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox GlovesImage 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   26
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   480
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   23
               Top             =   1320
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox Ring1Image 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   24
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   1680
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   21
               Top             =   1320
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox Ring2Image 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   22
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture10 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   1680
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   19
               Top             =   1920
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox BootsImage 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   20
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   555
               Left            =   1080
               ScaleHeight     =   35
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   35
               TabIndex        =   17
               Top             =   1320
               Visible         =   0   'False
               Width           =   555
               Begin VB.PictureBox LegsImage 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Segoe UI"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   495
                  Left            =   15
                  ScaleHeight     =   495
                  ScaleWidth      =   495
                  TabIndex        =   18
                  Top             =   15
                  Width           =   495
               End
            End
            Begin VB.PictureBox picItems 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   2.25000e5
               Left            =   2400
               Picture         =   "frmMirage.frx":21BD0
               ScaleHeight     =   2.23636e5
               ScaleMode       =   0  'User
               ScaleWidth      =   477.091
               TabIndex        =   16
               Top             =   2760
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.PictureBox picGuildAdmin 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   2385
            TabIndex        =   35
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
            Begin VB.TextBox txtAccess 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   720
               MaxLength       =   2
               TabIndex        =   41
               Top             =   585
               Width           =   1575
            End
            Begin VB.TextBox txtName 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
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
               Left            =   720
               TabIndex        =   40
               Top             =   345
               Width           =   1575
            End
            Begin VB.CommandButton cmdTrainee 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Recruter"
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
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   975
               Width           =   1815
            End
            Begin VB.CommandButton cmdMember 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Recruter (comme recruteur)"
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
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   1305
               Width           =   1815
            End
            Begin VB.CommandButton cmdDisown 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Faire quitter la Guilde"
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
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   1650
               Width           =   1815
            End
            Begin VB.CommandButton cmdAccess 
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               Caption         =   "Changer l'Access"
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
               Left            =   420
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1980
               Width           =   1815
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Access:"
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
               Left            =   150
               TabIndex        =   43
               Top             =   615
               Width           =   465
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nom:"
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
               Left            =   180
               TabIndex        =   42
               Top             =   360
               Width           =   345
            End
         End
         Begin VB.PictureBox picGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2505
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   2385
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   2385
            Begin VB.Label lblRank 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rank"
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
               Left            =   1425
               TabIndex        =   33
               Top             =   975
               Width           =   1080
            End
            Begin VB.Label lblGuild 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Guild"
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
               Left            =   1425
               TabIndex        =   32
               Top             =   660
               Width           =   1065
            End
            Begin VB.Label Label17 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Votre access:"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   480
               TabIndex        =   31
               Top             =   960
               Width           =   825
            End
            Begin VB.Label Label16 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Nom de la Guilde:"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   240
               TabIndex        =   30
               Top             =   645
               Width           =   1050
            End
            Begin VB.Label cmdLeave 
               BackStyle       =   0  'Transparent
               Caption         =   "Quitter la Guilde"
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
               Left            =   720
               TabIndex        =   34
               Top             =   2280
               Width           =   1110
            End
         End
         Begin VB.Label lblmaskinvferm 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   91
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblmaskinvmin 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2040
            TabIndex        =   90
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblmaskinv 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   0
            MousePointer    =   5  'Size
            TabIndex        =   61
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label ffermer 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   87
            Top             =   0
            Width           =   375
         End
         Begin VB.Label freduire 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   1920
            TabIndex        =   86
            Top             =   0
            Width           =   375
         End
         Begin VB.Image Image3 
            Height          =   2985
            Left            =   0
            Picture         =   "frmMirage.frx":181512
            Top             =   0
            Width           =   2595
         End
      End
      Begin VB.Frame picParty 
         BackColor       =   &H00004080&
         BorderStyle     =   0  'None
         Height          =   2985
         Left            =   120
         TabIndex        =   219
         Top             =   4440
         Visible         =   0   'False
         Width           =   2595
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   0
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   232
            Top             =   600
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   233
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   0
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   230
            Top             =   800
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   0
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   231
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox Picture15 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            Picture         =   "frmMirage.frx":19A98C
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   229
            Top             =   2400
            Width           =   270
         End
         Begin VB.PictureBox Picture16 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   960
            Picture         =   "frmMirage.frx":19AC24
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   228
            Top             =   2400
            Width           =   270
         End
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   1
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   226
            Top             =   1275
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   227
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   1
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   224
            Top             =   1485
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   1
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   225
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPLife 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   2
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   222
            Top             =   1995
            Width           =   2175
            Begin VB.Shape shpPPLife 
               BackColor       =   &H0000C000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   2
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPLife 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PV : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   223
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.PictureBox backPPMana 
            BackColor       =   &H00FF8080&
            BorderStyle     =   0  'None
            Height          =   170
            Index           =   2
            Left            =   240
            ScaleHeight     =   165
            ScaleWidth      =   2175
            TabIndex        =   220
            Top             =   2205
            Width           =   2175
            Begin VB.Shape shpPPMana 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderStyle     =   0  'Transparent
               Height          =   165
               Index           =   2
               Left            =   0
               Top             =   0
               Width           =   2175
            End
            Begin VB.Label lblPPMana 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "PM : "
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   5.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000005&
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   221
               Top             =   0
               Width           =   2175
            End
         End
         Begin VB.Image Image5 
            Height          =   2985
            Left            =   0
            Picture         =   "frmMirage.frx":19AEAF
            Top             =   0
            Width           =   2595
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   0
            MousePointer    =   5  'Size
            TabIndex        =   240
            Top             =   0
            Width           =   2655
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2280
            TabIndex        =   239
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rejoindre/Quitter le groupe"
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
            Left            =   0
            TabIndex        =   238
            Top             =   2760
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   240
            TabIndex        =   237
            Top             =   400
            Width           =   2175
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   2040
            TabIndex        =   236
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Index           =   1
            Left            =   240
            TabIndex        =   235
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblPPName 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Index           =   2
            Left            =   240
            TabIndex        =   234
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   6120
         Top             =   0
      End
      Begin VB.PictureBox picquete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4290
         Left            =   480
         Picture         =   "frmMirage.frx":1B4329
         ScaleHeight     =   286
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   92
         Top             =   960
         Visible         =   0   'False
         Width           =   3345
         Begin VB.TextBox quetetxt 
            Appearance      =   0  'Flat
            Height          =   3015
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label artquete 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   1440
            TabIndex        =   94
            Top             =   3840
            Width           =   1845
         End
         Begin VB.Label qf 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3045
            TabIndex        =   97
            Top             =   0
            Width           =   285
         End
         Begin VB.Label av 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   165
            Left            =   1440
            TabIndex        =   96
            Top             =   2040
            Width           =   45
         End
         Begin VB.Label qt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   240
            TabIndex        =   95
            Top             =   3600
            Width           =   1020
         End
      End
      Begin VB.Frame fra_info 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   5250
         Left            =   4050
         TabIndex        =   1
         Top             =   960
         Width           =   4500
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   2520
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   121
            Top             =   2640
            Width           =   555
            Begin VB.PictureBox PetImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   122
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1920
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   83
            Top             =   960
            Width           =   555
            Begin VB.PictureBox HelmetImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   84
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   2520
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   81
            Top             =   1560
            Width           =   555
            Begin VB.PictureBox ShieldImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   82
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1920
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   79
            Top             =   2640
            Width           =   555
            Begin VB.PictureBox ArmorImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   80
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1320
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   77
            Top             =   1560
            Width           =   555
            Begin VB.PictureBox WeaponImage 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   78
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.PictureBox Picsprt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   1920
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   35
            TabIndex        =   75
            Top             =   1560
            Width           =   555
            Begin VB.PictureBox Picsprts 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   15
               ScaleHeight     =   495
               ScaleWidth      =   495
               TabIndex        =   76
               Top             =   15
               Width           =   495
            End
         End
         Begin VB.Label lblMetierApl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Metier"
            Height          =   255
            Left            =   2880
            TabIndex        =   246
            Top             =   5040
            Width           =   735
         End
         Begin VB.Label ifermer 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   4200
            TabIndex        =   89
            Top             =   0
            Width           =   375
         End
         Begin VB.Label ireduire 
            BackStyle       =   0  'Transparent
            Caption         =   "                                   "
            Height          =   375
            Left            =   3840
            TabIndex        =   88
            Top             =   0
            Width           =   375
         End
         Begin VB.Label fermer 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   3720
            TabIndex        =   85
            Top             =   5040
            Width           =   855
         End
         Begin VB.Label AddDef 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   2640
            TabIndex        =   14
            Top             =   3930
            Width           =   165
         End
         Begin VB.Label AddMagi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   13
            Top             =   4185
            Width           =   165
         End
         Begin VB.Label AddSpeed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   12
            Top             =   4425
            Width           =   165
         End
         Begin VB.Label AddStr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2640
            TabIndex        =   11
            Top             =   3690
            Width           =   165
         End
         Begin VB.Label lblSPEED 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   10
            ToolTipText     =   "Points permettant d'augmenter vos chances d'esquive"
            Top             =   4455
            Width           =   270
         End
         Begin VB.Label lblMAGI 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            ToolTipText     =   "Points permettant d'augmenter vos sorts disponibles "
            Top             =   4215
            Width           =   270
         End
         Begin VB.Label lblDEF 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   1560
            TabIndex        =   8
            ToolTipText     =   "Points permettant d'augmenter votre résistance et vos chances de bloquer"
            Top             =   3960
            Width           =   270
         End
         Begin VB.Label lblSTR 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1560
            TabIndex        =   7
            ToolTipText     =   "Points permettant d'augmenter vos dégâts et vos chances de coup critique"
            Top             =   3720
            Width           =   270
         End
         Begin VB.Label lblPoints 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   300
            Left            =   1680
            TabIndex        =   6
            Top             =   4740
            Width           =   435
         End
         Begin VB.Label lblLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   180
            Left            =   1320
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label maclasse 
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   210
            Left            =   2400
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label monnom 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CC"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   600
            TabIndex        =   3
            Top             =   480
            Width           =   240
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            MousePointer    =   5  'Size
            Top             =   0
            Width           =   4545
         End
         Begin VB.Image Image1 
            Height          =   5250
            Left            =   0
            Picture         =   "frmMirage.frx":1E322B
            Top             =   0
            Width           =   4500
         End
      End
      Begin VB.PictureBox xp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   73
         Top             =   600
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lexp 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape sexp 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            FillColor       =   &H000000FF&
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox mana 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   71
         Top             =   360
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lmana 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00CB884B&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape smana 
            BackColor       =   &H00CB884B&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox vie 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   10320
         ScaleHeight     =   180
         ScaleWidth      =   1425
         TabIndex        =   69
         Top             =   120
         Visible         =   0   'False
         Width           =   1425
         Begin VB.Label lvie 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   1425
         End
         Begin VB.Shape svie 
            BackColor       =   &H0000C000&
            BackStyle       =   1  'Opaque
            BorderStyle     =   0  'Transparent
            Height          =   180
            Left            =   0
            Top             =   0
            Width           =   1425
         End
      End
      Begin VB.PictureBox ObjNm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4440
         ScaleHeight     =   255
         ScaleWidth      =   1575
         TabIndex        =   66
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Label OName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   195
            Left            =   120
            TabIndex        =   67
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Timer Timer1 
         Left            =   7320
         Top             =   0
      End
      Begin VB.Timer tmrSnowDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8760
         Top             =   0
      End
      Begin VB.Timer tmrRainDrop 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   8280
         Top             =   0
      End
      Begin VB.PictureBox ScreenShot 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   9240
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox txtQ 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         Picture         =   "frmMirage.frx":2300E5
         ScaleHeight     =   1545
         ScaleWidth      =   9510
         TabIndex        =   56
         Top             =   7440
         Visible         =   0   'False
         Width           =   9540
         Begin VB.TextBox TxtQ2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   1065
            Left            =   158
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Text            =   "frmMirage.frx":25FF37
            Top             =   180
            Width           =   9160
         End
         Begin VB.Label OK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "OK"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   9120
            TabIndex        =   58
            Top             =   1360
            Width           =   495
         End
      End
   End
   Begin VB.Label lbltimeQuete 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7800
      TabIndex        =   252
      Top             =   9185
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label menu_quete 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   9240
      TabIndex        =   142
      ToolTipText     =   "Quetes"
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label menu_quit 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11640
      TabIndex        =   119
      ToolTipText     =   "Quitter"
      Top             =   9240
      Width           =   345
   End
   Begin VB.Label menu_equ 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8760
      TabIndex        =   118
      ToolTipText     =   "Equipements"
      Top             =   9240
      Width           =   345
   End
   Begin VB.Label menu_guild 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9840
      TabIndex        =   117
      ToolTipText     =   "Guilde"
      Top             =   9240
      Width           =   405
   End
   Begin VB.Label menu_opt 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   11160
      TabIndex        =   116
      ToolTipText     =   "Options"
      Top             =   9240
      Width           =   420
   End
   Begin VB.Label menu_tchat 
      BackStyle       =   0  'Transparent
      Height          =   540
      Left            =   10320
      TabIndex        =   115
      ToolTipText     =   "Tchat"
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label menu_who 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10560
      TabIndex        =   114
      ToolTipText     =   "Qui est en ligne ?"
      Top             =   9240
      Width           =   540
   End
   Begin VB.Label menu_sort 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8280
      TabIndex        =   113
      ToolTipText     =   "Sorts"
      Top             =   9240
      Width           =   465
   End
   Begin VB.Label menu_inv 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7920
      TabIndex        =   112
      ToolTipText     =   "Inventaire"
      Top             =   9240
      Width           =   315
   End
   Begin VB.Image Interface 
      Height          =   900
      Left            =   0
      Picture         =   "frmMirage.frx":25FF3D
      Top             =   9120
      Width           =   12000
   End
   Begin WMPLibCtl.WindowsMediaPlayer Mediaplayer 
      Height          =   720
      Left            =   12360
      TabIndex        =   65
      Top             =   4560
      Width           =   480
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "invisible"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   847
      _cy             =   1270
   End
End
Attribute VB_Name = "frmMirage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'1024X768

Private SpellMemorized As Long
Public DragImg As Long
Public DragX As Long
Public DragY As Long
Private OldPCX As Long
Private OldPCY As Long
Private twippx As Long
Private twippy As Long
Private Lon As Long
Private Hau As Long

Private Sub AddDef_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 1 & END_CHAR)
End Sub

Private Sub AddMagi_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 2 & END_CHAR)
End Sub

Private Sub AddSpeed_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 3 & END_CHAR)
End Sub

Private Sub AddStr_Click()
    Call SendData("usestatpoint" & SEP_CHAR & 0 & END_CHAR)
End Sub

Private Sub artquete_Click()
    Player(MyIndex).QueteEnCour = 0
    Accepter = False
    Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
    frmMirage.picquete.Visible = False
    If quetetimersec.Enabled Then
        quetetimersec.Enabled = False
        tmpsquete.Visible = False
    End If
End Sub

Private Sub cbtr1_Change()

End Sub

Private Sub cbth_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtb_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtg_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtd_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbta_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtra_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtc_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtac_keypress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cbtr_keypress(Index As Integer, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub chkLowEffect_Click()
    WriteINI "CONFIG", "LowEffect", chkLowEffect.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknobj_Click()
    WriteINI "CONFIG", "NomObjet", chknobj.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chksound_Click()
    WriteINI "CONFIG", "Sound", chksound.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkbubblebar_Click()
    WriteINI "CONFIG", "SpeechBubbles", chkbubblebar.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcbar_Click()
    WriteINI "CONFIG", "NpcBar", chknpcbar.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcdamage_Click()
    WriteINI "CONFIG", "NPCDamage", chknpcdamage.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chknpcname_Click()
    WriteINI "CONFIG", "NPCName", chknpcname.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayerbar_Click()
    WriteINI "CONFIG", "PlayerBar", chkplayerbar.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayerdamage_Click()
    WriteINI "CONFIG", "PlayerDamage", chkplayerdamage.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkAutoScroll_Click()
    WriteINI "CONFIG", "AutoScroll", chkAutoScroll.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkplayername_Click()
    WriteINI "CONFIG", "PlayerName", chkplayername.Value, App.Path & "\Config\Account.ini"
End Sub

Private Sub chkmusic_Click()
    WriteINI "CONFIG", "Music", chkmusic.Value, App.Path & "\Config\Account.ini"
    If MyIndex <= 0 Then Exit Sub
    Call PlayMidi(Trim$(Map(GetPlayerMap(MyIndex)).Music))
End Sub

Private Sub cmdLeave_Click()
Dim Packet As String
    Packet = "GUILDLEAVE" & END_CHAR
    Call SendData(Packet)
    lblGuild.Caption = vbNullString
    lblRank.Caption = 0
End Sub

Private Sub cmdMember_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "GUILDMEMBER" & SEP_CHAR & txtName.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub CmdoptTouche_Click()
Dim i As Byte, n As Byte
cbth.Clear
cbtb.Clear
cbtg.Clear
cbtd.Clear
cbta.Clear
cbtc.Clear
cbtra.Clear
cbtac.Clear
For i = 0 To 13
    cbtr(i).Clear
Next i

For i = 0 To TCHMAX
        cbth.AddItem optTouche(i).nom, i
        cbtb.AddItem optTouche(i).nom, i
        cbtg.AddItem optTouche(i).nom, i
        cbtd.AddItem optTouche(i).nom, i
        cbta.AddItem optTouche(i).nom, i
        cbtc.AddItem optTouche(i).nom, i
        cbtra.AddItem optTouche(i).nom, i
        cbtac.AddItem optTouche(i).nom, i
        For n = 0 To 13
            cbtr(n).AddItem optTouche(i).nom, i
        Next n
Next i

    cbth.ListIndex = CByte(Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini")))
    cbth.Text = optTouche(CByte(Val(ReadINI("TJEU", "haut", App.Path & "\Config\Option.ini")))).nom
    cbtb.ListIndex = CByte(Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini")))
    cbtb.Text = optTouche(CByte(Val(ReadINI("TJEU", "bas", App.Path & "\Config\Option.ini")))).nom
    cbtg.ListIndex = CByte(Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini")))
    cbtg.Text = optTouche(CByte(Val(ReadINI("TJEU", "gauche", App.Path & "\Config\Option.ini")))).nom
    cbtd.ListIndex = CByte(Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini")))
    cbtd.Text = optTouche(CByte(Val(ReadINI("TJEU", "droite", App.Path & "\Config\Option.ini")))).nom
    cbta.ListIndex = CByte(Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini")))
    cbta.Text = optTouche(CByte(Val(ReadINI("TJEU", "attaque", App.Path & "\Config\Option.ini")))).nom
    cbtc.ListIndex = CByte(Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini")))
    cbtc.Text = optTouche(CByte(Val(ReadINI("TJEU", "courir", App.Path & "\Config\Option.ini")))).nom
    cbtra.ListIndex = CByte(Val(ReadINI("TJEU", "ramasser", App.Path & "\Config\Option.ini")))
    cbtra.Text = optTouche(CByte(Val(ReadINI("TJEU", "ramasser", App.Path & "\Config\Option.ini")))).nom
    cbtac.ListIndex = CByte(Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini")))
    cbtac.Text = optTouche(CByte(Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini")))).nom
    For i = 0 To 13
        cbtr(i).ListIndex = CByte(Val(ReadINI("TRAC", "rac" & (i + 1), App.Path & "\Config\Option.ini")))
        cbtr(i).Text = optTouche(CByte(Val(ReadINI("TRAC", "rac" & (i + 1), App.Path & "\Config\Option.ini")))).nom
    Next i
pictTouche.Visible = True
End Sub

Private Sub cmdOTA_Click()
    pictTouche.Visible = False
End Sub

Private Sub cmdOTO_Click()
Dim i As Byte
    Call WriteINI("TJEU", "haut", cbth.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "bas", cbtb.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "gauche", cbtg.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "droite", cbtd.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "attaque", cbta.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "courir", cbtc.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "ramasser", cbtra.ListIndex, App.Path & "\Config\Option.ini")
    Call WriteINI("TJEU", "action", cbtac.ListIndex, App.Path & "\Config\Option.ini")
    For i = 0 To 13
        Call WriteINI("TRAC", "raci", cbtr(i).ListIndex, App.Path & "\Config\Option.ini")
    Next i
    pictTouche.Visible = False
End Sub

Private Sub Command1_Click()
picOptions.Visible = False
Call InitAccountOpt
End Sub

Private Sub Command2_Click()
Call Form_Load
End Sub

Private Sub fermer_Click()
fra_info.Visible = False
End Sub

Private Sub ffermer_Click()
fra_fenetre.Visible = False
End Sub

Private Sub Form_GotFocus()
Picsprt.height = 48
Picsprts.height = 48
On Error Resume Next
txtMyTextBox.SetFocus

End Sub

Private Sub Form_Load()
Dim i As Long, x As Integer
Dim Ending As String
Dim Qq As Long

        
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExiste(Rep_Theme & "\Jeu\Text" & Ending) Then txtQ.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\text" & Ending)
        If FileExiste(Rep_Theme & "\info" & Ending) Then frmMirage.Picture = LoadPNG(App.Path & Rep_Theme & "\info" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\Info" & Ending) Then Image1.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Info" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\inventaire" & Ending) Then Image3.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\inventaire" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\Carte" & Ending) Then imgcarte.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\Carte" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\quitter" & Ending) Then PicMenuQuitter.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\quitter" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\quete" & Ending) Then picquete.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\quete" & Ending)
        If FileExiste(Rep_Theme & "\Jeu\metier" & Ending) Then pictMetier.Picture = LoadPNG(App.Path & Rep_Theme & "\Jeu\metier" & Ending)
        
    Next i

    Call netbook_change
    
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX
    svie.FillColor = RGB(208, 11, 0)
    smana.FillColor = RGB(208, 11, 0)
    
    'If frmMainMenu.chk_fullscreen.value = Checked Then
        'If (Screen.Height / Screen.TwipsPerPixelY) >= 758 Then txtMyTextBox.Top = 567
        'frmMirage.Height = Screen.Height / Screen.TwipsPerPixelY
        'frmMirage.Width = Screen.Width / Screen.TwipsPerPixelX
        'picScreen.Height = Screen.Height / Screen.TwipsPerPixelY
        'picScreen.Width = Screen.Width / Screen.TwipsPerPixelX
    'End If

    monnom.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    monnom.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    maclasse.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    maclasse.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    txtMyTextBox.Font = ReadINI("POLICE", "PoliceChat", (App.Path & "\Config\Ecriture.ini"))
    
    fra_info.Visible = False
    fra_fenetre.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If GettingMap Then Cancel = True Else Call GameDestroy
End Sub

Private Sub freduire_Click()
    If fra_fenetre.height >= 2985 / 15 Then fra_fenetre.height = 315 / 15 Else fra_fenetre.height = 2985 / 15
End Sub

Private Sub ifermer_Click()
    fra_info.Visible = False
End Sub

Private Sub Image3_Click()
    fra_fenetre.Visible = False
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 3
DragX = x
DragY = y
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 3 Then fra_info.Top = fra_info.Top + ((y / twippy) - (DragY / twippy)): fra_info.Left = fra_info.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub ireduire_Click()
    If fra_info.height >= 350 Then fra_info.height = 315 / twippy Else fra_info.height = 350
End Sub

Private Sub Label19_Click()
    If picParty.height >= 2985 / twippy Then picParty.height = 315 / twippy Else picParty.height = 2985 / twippy
End Sub

Private Sub Label27_Click()
    PicMenuQuitter.Visible = False
End Sub

Private Sub Label3_Click()
    If Player(MyIndex).PartyIndex > 0 Then SendLeaveParty: picParty.Visible = False Else picParty.Visible = False
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 6
DragX = x
DragY = y
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 6 Then picParty.Move picParty.Left + ((x / twippx) - (DragX / twippx)), picParty.Top + ((y / twippy) - (DragY / twippy))
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub Label6_Click()
    Call SendData("getstats" & END_CHAR)
End Sub

Private Sub Label8_Click()
    picParty.Visible = False
End Sub

Private Sub lblCdP_Click()
    Call SendData("CHANGECHAR" & END_CHAR)
    frmMirage.Visible = False
    frmMainMenu.Visible = True
    frmMainMenu.fraPers.Visible = True
    frmsplash.Visible = False
    PicMenuQuitter.Visible = False
End Sub

Private Sub lblDeco_Click()
Dim i As Integer
    Call SendData("CHANGECHAR" & END_CHAR)
    InGame = False
    deco = True
    Sleep 2000
    PicMenuQuitter.Visible = False
    frmMainMenu.Visible = True
    frmMainMenu.fraLogin.Visible = True
    frmMainMenu.fraPers.Visible = False
    frmMirage.Visible = False
    frmMirage.Socket.Close
    frmMirage.Socket.Connect
End Sub

Private Sub lblendmetier_Click()
pictMetier.Visible = False
End Sub

Private Sub lblmaskinvferm_Click()
fra_fenetre.Visible = False
End Sub

Private Sub lblmaskinvmin_Click()
    If fra_fenetre.height >= 2985 / twippy Then fra_fenetre.height = 315 / twippy Else fra_fenetre.height = 2985 / twippy
End Sub

Private Sub lblmaskinv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 2
DragX = x
DragY = y
End Sub

Private Sub lblmaskinv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 2 Then fra_fenetre.Top = fra_fenetre.Top + ((y / twippy) - (DragY / twippy)): fra_fenetre.Left = fra_fenetre.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub lblmaskinv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub



Private Sub lblmaskmenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 1
DragX = x
DragY = y
End Sub



Private Sub lblmaskmenu_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub lblMetierApl_Click()
    Call SendData("playermetier" & END_CHAR)
End Sub

Private Sub lblmetierEnd_Click()
    pictMetier.Visible = False
End Sub

Private Sub lblOublierMetier_Click()
    Call SendData("playermetieroublie" & END_CHAR)
    pictMetier.Visible = False
End Sub

Private Sub lblPoints_Change()
    With frmMirage
        If GetPlayerPOINTS(MyIndex) > 0 Then
            .AddStr.Visible = True
            .AddDef.Visible = True
            .AddSpeed.Visible = True
            .AddMagi.Visible = True
        Else
            .AddStr.Visible = False
            .AddDef.Visible = False
            .AddSpeed.Visible = False
            .AddMagi.Visible = False
        End If
    End With
End Sub

Private Sub lblQuitter_Click()
    Call GameDestroy
End Sub

Private Sub lstOnline_DblClick()
    Call SendData("playerchat" & SEP_CHAR & Trim$(lstOnline.Text) & END_CHAR)
End Sub

Private Sub menu_equ_Click()

If picOptions.Visible = True Then picOptions.Visible = False
If fra_info.Visible = True Then
    fra_info.Visible = False
Else
    fra_fenetre.Visible = False
    Call ClearPic
    Call UpdateVisInv
    PrepareSprite (Player(MyIndex).Sprite)
    fra_info.Visible = True
    Picsprt.height = (48 + 4) * twippy
    Picsprts.height = (48)
    If Picsprts.height <= 32 Then Picture5.Top = 2160 Else Picture5.Top = 2640
    Call AffSurfPic(DD_SpriteSurf(Player(MyIndex).Sprite), Picsprts, 0, 0)
    'Call BitBlt(Picsprts.hDC, 0, 0, PIC_X, PIC_Y * PIC_NPC1, Picturesprite.hDC, 3 * PIC_X, Val(Player(MyIndex).Sprite) * (PIC_Y * PIC_NPC1), SRCCOPY)
End If

End Sub

Private Sub menu_guild_Click()

If picOptions.Visible = True Then picOptions.Visible = False
' Set Their Guild Name and Their Rank
If fra_fenetre.Visible = True And picGuild.Visible = True Then fra_fenetre.Visible = False Else fra_fenetre.Visible = True
picParty.Visible = (fra_fenetre.Visible And (Player(MyIndex).PartyIndex > 0))
Label3.Visible = picParty.Visible
If picParty.Visible Then
    Dim i As Integer, C As Byte
    If lblPPName(0).Tag <= lblPPName(2).Tag Or lblPPName(2).Caption <> vbNullString Then
        For i = (Val(lblPPName(2).Tag) + 1) To MAX_PLAYERS
            If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C < 3 And i <> MyIndex Then
                C = C + 1
                lblPPName(C - 1).Tag = i
            End If
        Next
        For i = 0 To 2
            lblPPName(i).Visible = (i < C)
            backPPLife(i).Visible = lblPPName(i).Visible
            backPPMana(i).Visible = lblPPName(i).Visible
            If lblPPName(i).Visible Then
                lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
            End If
        Next
    End If
End If
Call ClearPic
frmMirage.lblGuild.Caption = GetPlayerGuild(MyIndex)
frmMirage.lblRank.Caption = GetPlayerGuildAccess(MyIndex)
picGuild.Visible = True
End Sub

Private Sub menu_inv_Click()
If picOptions.Visible = True Then picOptions.Visible = False
If fra_fenetre.Visible = True And picInv3.Visible = True Then fra_fenetre.Visible = False Else fra_fenetre.Visible = True
Call UpdateVisInv
Call ClearPic
picInv3.Visible = True
End Sub

Private Sub menu_opt_Click()
    fra_info.Visible = False
    If fra_fenetre.Visible = True Then fra_fenetre.Visible = False
    If picquete.Visible = True Then picquete.Visible = False
    
    
    If picOptions.Visible = False Then
        picOptions.Visible = True
    Else
        picOptions.Visible = False
    End If
    

End Sub

Private Sub menu_quete_Click()

If picOptions.Visible = True Then picOptions.Visible = False


If frmMirage.picquete.Visible = True Then
    frmMirage.picquete.Visible = False
Else

    If Player(MyIndex).QueteEnCour > 0 Then
        Call ClearPic
        fra_fenetre.Visible = False
        frmMirage.picquete.Visible = True
        frmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description
    Else
        Call ClearPic
        fra_fenetre.Visible = False
        frmMirage.picquete.Visible = True
        frmMirage.quetetxt.Text = "Pas de quête en cours..."
    End If
    
End If
End Sub

Private Sub menu_quit_Click()
'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
'Call GameDestroy
If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
End Sub

Private Sub menu_sort_Click()
If picOptions.Visible = True Then picOptions.Visible = False
If fra_fenetre.Visible = True And picPlayerSpells.Visible = True Then
    fra_fenetre.Visible = False
Else
    fra_fenetre.Visible = True
    picPlayerSpells.Visible = True
End If
Call ClearPic
Call SendData("spells" & END_CHAR)
End Sub

Private Sub menu_tchat_Click()
Dim i As Long
If picOptions.Visible = True Then picOptions.Visible = False
fra_info.Visible = False
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
    If MouseDownX = GetPlayerX(i) And MouseDownY = GetPlayerY(i) Then
    Call SendData("playerchat" & SEP_CHAR & GetPlayerName(i) & END_CHAR): Exit Sub
    Else
    MsgBox ("Vous devez sélectionner un joueur")
    End If
    End If
Next i

End Sub

Private Sub menu_who_Click()
If picOptions.Visible = True Then picOptions.Visible = False
    If fra_fenetre.Visible = True And picWhosOnline.Visible = True Then fra_fenetre.Visible = False Else fra_fenetre.Visible = True

    Call SendOnlineList
    
    Call ClearPic
    picWhosOnline.Visible = True
End Sub

Private Sub OK_Click()
Dim i As Long
Dim msgb As String

If Player(MyIndex).QueteEnCour > 0 And Accepter = False Then
    msgb = MsgBox("Voulez-vous faire la quête proposée ?", vbYesNo, "Quete")
        If msgb = vbYes Then
            Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
            Accepter = True
        Else
            Player(MyIndex).QueteEnCour = 0
            Call SendData("DEMAREQUETE" & SEP_CHAR & Player(MyIndex).QueteEnCour & END_CHAR)
            Accepter = False
        End If
End If
txtQ.Visible = False
End Sub

Private Sub picInv_DblClick(Index As Integer)
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Or Player(MyIndex).Inv(Inventory).num > MAX_ITEMS Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
    End If
Next d
Call UpdateVisInv
End Sub

Private Sub picInv_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Inventory = Index + 1
    frmMirage.SelectedItem.Top = frmMirage.picInv(Inventory - 1).Top - 1
    frmMirage.SelectedItem.Left = frmMirage.picInv(Inventory - 1).Left - 1
    
    If Button = 1 Then
        Call UpdateVisInv
    ElseIf Button = 2 Then
        If Player(MyIndex).Inv(Inventory).num <= 0 Or Player(MyIndex).Inv(Inventory).num > MAX_ITEMS Then
            dragAndDropT = 0
            dragAndDrop = 0
            IDAD.Visible = False
        Else
            If dragAndDrop = Inventory Then
                dragAndDrop = 0
                dragAndDropT = 0
                IDAD.Visible = False
            Else
                dragAndDrop = Inventory
                dragAndDropT = 2
                IDAD.Top = frmMirage.picInv(Inventory - 1).Top - 1
                IDAD.Left = frmMirage.picInv(Inventory - 1).Left - 1
                IDAD.Visible = True
            End If
        End If
    ElseIf Button = 3 Then
        Call DropItems
    End If
End Sub

Private Sub picInv_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim d As Long
d = Index

    If Player(MyIndex).Inv(d + 1).num > 0 Then
        
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY And Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
            itmDesc.height = 17
            'itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            'itmDesc.Left = fra_fenetre.Left
            If netbook = True Then
                frmMirage.itmDesc.Left = frmMirage.fra_fenetre.Left - frmMirage.itmDesc.Width
                frmMirage.itmDesc.Top = frmMirage.picScreen.height - frmMirage.itmDesc.height - 10
            Else
                frmMirage.itmDesc.Left = frmMirage.picScreen.Width - frmMirage.itmDesc.Width - 30
                frmMirage.itmDesc.Top = frmMirage.fra_fenetre.Top - frmMirage.itmDesc.height
            End If
        ElseIf Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) = vbNullString Then
            itmDesc.height = 161
            'itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            'itmDesc.Left = fra_fenetre.Left
            If netbook = True Then
                frmMirage.itmDesc.Left = frmMirage.fra_fenetre.Left - frmMirage.itmDesc.Width
                frmMirage.itmDesc.Top = frmMirage.picScreen.height - frmMirage.itmDesc.height - 10
            Else
                frmMirage.itmDesc.Left = frmMirage.picScreen.Width - frmMirage.itmDesc.Width - 30
                frmMirage.itmDesc.Top = frmMirage.fra_fenetre.Top - frmMirage.itmDesc.height
            End If
        ElseIf Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc) > vbNullString Then
            itmDesc.height = 249
            'itmDesc.Top = fra_fenetre.Top - itmDesc.Height
            'itmDesc.Left = fra_fenetre.Left
            If netbook = True Then
                frmMirage.itmDesc.Left = frmMirage.fra_fenetre.Left - frmMirage.itmDesc.Width
                frmMirage.itmDesc.Top = frmMirage.picScreen.height - frmMirage.itmDesc.height - 10
            Else
                frmMirage.itmDesc.Left = frmMirage.picScreen.Width - frmMirage.itmDesc.Width - 30
                frmMirage.itmDesc.Top = frmMirage.fra_fenetre.Top - frmMirage.itmDesc.height
            End If
        End If
                
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable <> 0 Then
            descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
            ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
            ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
            ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (worn)"
            ElseIf GetPlayerPetSlot(MyIndex) = d + 1 Then
                descName.Caption = Trim$(Pets(GetPlayerInvItemNum(MyIndex, d + 1)).nom) & " (worn)"
            ElseIf Item(GetPlayerInvItemNum(MyIndex, d + 1)).Empilable <> 0 Then
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
            Else
                descName.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).name)
            End If
        End If
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_PET Then
            descStr.Caption = Pets(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1).addForce & " Force"
            descDef.Caption = Pets(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1).addDefence & " Défense"
        Else
            descStr.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).StrReq & " Force"
            descDef.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).DefReq & " Défense"
        End If
        descSpeed.Caption = Item(GetPlayerInvItemNum(MyIndex, d + 1)).SpeedReq & " Vitesse"
        descHpMp.Caption = "PV: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddHP & " PM: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMP & " End: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSP
        descSD.Caption = "FOR: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddStr & " Def: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddDef
        descMS.Caption = "Magie: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddMagi & " Vitesse: " & Item(GetPlayerInvItemNum(MyIndex, d + 1)).AddSpeed
        If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1 <= 0 Then Usure.Caption = "Usure : Ind." Else Usure.Caption = "Usure : " & GetPlayerInvItemDur(MyIndex, d + 1) & "/" & Item(GetPlayerInvItemNum(MyIndex, d + 1)).Data1
        desc.Caption = Trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).desc)
        descName.ForeColor = Item(GetPlayerInvItemNum(MyIndex, d + 1)).NCoul
        itmDesc.Visible = True
    Else
        itmDesc.Visible = False
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim d As Long, i As Long
Dim ii As Long
Dim PX As Long
Dim PY As Long
Dim Cod As String
Dim tp As Long
    If ConOff = True Then Exit Sub

    Call CheckInput(0, KeyCode, Shift)
    
    If (frmMirage.txtMyTextBox.Visible = False) And (KeyCode = optTouche(CByte(Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini")))).Value) Then
        PX = 0
        PY = 0
        If Player(MyIndex).y - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_UP Then PX = 0: PY = -1
        End If
                
        If Player(MyIndex).y + 1 < MAX_MAPY + 1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_DOWN Then PX = 0: PY = 1
        End If
                
        If Player(MyIndex).x - 1 > -1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_LEFT Then PX = -1: PY = 0
        End If
        
        If Player(MyIndex).x + 1 < MAX_MAPX + 1 And PX = 0 And PY = 0 Then
            tp = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type
            If tp = TILE_TYPE_COFFRE Or tp = TILE_TYPE_PORTE_CODE And Player(MyIndex).Dir = DIR_RIGHT Then PX = 1: PY = 0
        End If
        
        If PX <> 0 Or PY <> 0 Then
        With Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY)
            If .String1 > vbNullString And TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = NO Then
                Dim Packet As String
                Cod = InputBox("Veuillez entre le mot de passe :", "Code")
                If Cod = .String1 Then
                    TempTile(GetPlayerX(MyIndex) + PX, GetPlayerY(MyIndex) + PY).DoorOpen = YES
                    Packet = "OUVRIRE" & SEP_CHAR & GetPlayerX(MyIndex) + PX & SEP_CHAR & GetPlayerY(MyIndex) + PY & END_CHAR
                    Call SendData(Packet)
                    If .Type = TILE_TYPE_COFFRE Then
                        i = FindOpenInvSlot(Val(.Data3))
                        If i > 0 Then
                            Call SetPlayerInvItemNum(MyIndex, i, Val(.Data3))
                            Call SetPlayerInvItemValue(MyIndex, i, GetPlayerInvItemValue(MyIndex, i) + 1)
                            Call SetPlayerInvItemDur(MyIndex, i, Item(Val(.Data3)).Data1)
                            Call UpdateVisInv
                            Packet = "ACOFFRE" & SEP_CHAR & i & SEP_CHAR & Val(.Data3) & SEP_CHAR & 1 & SEP_CHAR & Item(Val(.Data3)).Data1 & END_CHAR
                            Call SendData(Packet)
                        End If
                    End If
                Else
                    Call MsgBox("Mauvais code.", vbCritical)
                End If
            End If
        End With
        End If
        
        If GetPlayerY(MyIndex) - 1 > 0 And GetPlayerY(MyIndex) - 1 < MAX_MAPY Then
            With Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1)
            If .Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                If Trim$(.String1) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String1))
                If Trim$(.String2) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String2))
                If Trim$(.String3) <> vbNullString Then Call QueteMsg(MyIndex, "Il est marqué: " & Trim$(.String3))
                Exit Sub
            End If
            End With
        End If
    End If
    
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac1", App.Path & "\Config\Option.ini")))).Value Then Call useRac(0)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac2", App.Path & "\Config\Option.ini")))).Value Then Call useRac(1)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac3", App.Path & "\Config\Option.ini")))).Value Then Call useRac(2)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac4", App.Path & "\Config\Option.ini")))).Value Then Call useRac(3)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac5", App.Path & "\Config\Option.ini")))).Value Then Call useRac(4)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac6", App.Path & "\Config\Option.ini")))).Value Then Call useRac(5)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac7", App.Path & "\Config\Option.ini")))).Value Then Call useRac(6)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac8", App.Path & "\Config\Option.ini")))).Value Then Call useRac(7)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac9", App.Path & "\Config\Option.ini")))).Value Then Call useRac(8)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac10", App.Path & "\Config\Option.ini")))).Value Then Call useRac(9)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac11", App.Path & "\Config\Option.ini")))).Value Then Call useRac(10)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac12", App.Path & "\Config\Option.ini")))).Value Then Call useRac(11)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac13", App.Path & "\Config\Option.ini")))).Value Then Call useRac(12)
    If KeyCode = optTouche(CByte(Val(ReadINI("TRAC", "rac14", App.Path & "\Config\Option.ini")))).Value Then Call useRac(13)
    
    If KeyCode = vbKeyEscape Then
        If PicMenuQuitter.Visible Then PicMenuQuitter.Visible = False Else PicMenuQuitter.Visible = True
    End If

    ' The Guild Maker
    If KeyCode = vbKeyF5 Then
        If Player(MyIndex).Guildaccess > 1 Then Call ClearPic: fra_fenetre.Visible = True: frmMirage.picGuildAdmin.Visible = True Else Call ClearPic
    End If
    
    ' The Guild Creator
    If KeyCode = vbKeyF6 Then frmGuild.txtName = GetPlayerName(MyIndex): frmGuild.Show vbModeless, frmMirage
    
    'quete desc
    If KeyCode = vbKeyF7 Then
        If Player(MyIndex).QueteEnCour > 0 Then Call ClearPic: fra_fenetre.Visible = False: frmMirage.picquete.Visible = True: frmMirage.quetetxt.Text = quete(Player(MyIndex).QueteEnCour).description Else Call ClearPic
    End If
    
    If KeyCode = vbKeyF8 Then frmPlayerHelp.Show
    
    If KeyCode = vbKeyF9 Then If Player(MyIndex).Access > 0 Then frmadmin.Show
    
    If KeyCode = vbKeyInsert Then
        If SpellMemorized > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & SpellMemorized & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Vous ne pouvez lancer un sort en marchant.", BrightRed)
                End If
            End If
        Else
            Call AddText("Aucune magie mémoriser.", BrightRed)
        End If
    Else
        Call CheckInput(0, KeyCode, Shift)
    End If
    
    If KeyCode = vbKeyF11 Then
        ScreenShot.Picture = CaptureForm(frmMirage)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    ElseIf KeyCode = vbKeyF12 Then
        ScreenShot.Picture = CaptureArea(frmMirage, 8, 6, 634, 479)
        i = 0
        ii = 0
        Do
            If FileExiste("Screenshot" & i & ".bmp") = True Then i = i + 1 Else Call SavePicture(ScreenShot.Picture, App.Path & "\Screenshot" & i & ".bmp"): ii = 1
            DoEvents
            Sleep 1
        Loop Until ii = 1
    End If
    
    If KeyCode = vbKeyEnd Then
    d = GetPlayerDir(MyIndex)
    
        If Player(MyIndex).Moving = NO Then
            If Player(MyIndex).Dir = DIR_DOWN Then
                Call SetPlayerDir(MyIndex, DIR_LEFT)
                If d <> DIR_LEFT Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_LEFT Then
                Call SetPlayerDir(MyIndex, DIR_UP)
                If d <> DIR_UP Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_UP Then
                Call SetPlayerDir(MyIndex, DIR_RIGHT)
                If d <> DIR_RIGHT Then Call Sendplayerdir
            ElseIf Player(MyIndex).Dir = DIR_RIGHT Then
                Call SetPlayerDir(MyIndex, DIR_DOWN)
                If d <> DIR_DOWN Then Call Sendplayerdir
            End If
        End If
    End If
End Sub

Private Sub PicOptions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SOffsetX = x
    SOffsetY = y
End Sub

Private Sub PicOptions_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MovePicture(frmMirage.picOptions, Button, Shift, x, y)
End Sub

Private Sub picquete_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 5
DragX = x
DragY = y
End Sub

Private Sub picquete_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 5 Then DoEvents: If DragImg = 5 Then picquete.Top = picquete.Top + ((y / twippy) - (DragY / twippy)): picquete.Left = picquete.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub picquete_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub picRac_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Q As Long
Dim Qq As Long
Dim d As Byte
    If Button = 1 Then
        Call useRac(Index)
    End If
    If Button = 2 Then
        If dragAndDrop > 0 Then
            rac(Index, 0) = dragAndDrop
            rac(Index, 1) = dragAndDropT
        End If
        Call saveRac
    End If
    dragAndDropT = 0
    dragAndDrop = 0
    SDAD.Visible = False
    IDAD.Visible = False
End Sub

Private Sub picScreen_GotFocus()
On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Call PlayerSearch(Button, Shift, (x + NewPlayerPicX), (y + NewPlayerPicY))
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
CurX = ((x + NewPlayerPicX) \ 32)
CurY = ((y + NewPlayerPicY) \ 32)
PotX = x
PotY = y

If CurX <> OldPCX Or CurY <> OldPCY Then Call CaseChange(CurX, CurY): OldPCX = CurX: OldPCY = CurY
itmDesc.Visible = False
End Sub

Private Sub picspell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If Player(MyIndex).Spell(Index + 1) > 0 Then
            If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
                If Player(MyIndex).Moving = 0 Then
                    Call SendData("cast" & SEP_CHAR & Index + 1 & END_CHAR)
                    Player(MyIndex).Attacking = 1
                    Player(MyIndex).AttackTimer = GetTickCount
                    Player(MyIndex).CastedSpell = YES
                Else
                    Call AddText("Vous ne pouvez lancer un sort en marchant.", BrightRed)
                End If
            End If
        Else
            Call AddText("Aucuns sort ici.", BrightRed)
        End If
    End If
    If Button = 2 Then
        If Player(MyIndex).Spell(Index + 1) > 0 Then
            If dragAndDrop = Index + 1 Then
                dragAndDrop = 0
                dragAndDropT = 0
                SDAD.Visible = False
            Else
                dragAndDrop = Index + 1
                dragAndDropT = 1
                SDAD.Top = picspell(Index).Top - 1
                SDAD.Left = picspell(Index).Left - 1
                SDAD.Visible = True
            End If
        Else
            dragAndDropT = 0
            dragAndDrop = 0
            SDAD.Visible = False
        End If
    End If
End Sub

Private Sub Picture15_Click()
    If Player(MyIndex).PartyIndex > 0 Then
        Dim i As Integer, C As Byte
        If lblPPName(0).Tag <= lblPPName(2).Tag And lblPPName(2).Caption <> vbNullString Then
            For i = (Val(lblPPName(2).Tag) + 1) To MAX_PLAYERS
                If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C < 3 And i <> MyIndex Then
                    C = C + 1
                    lblPPName(C - 1).Tag = i
                End If
            Next
            For i = 0 To 2
                lblPPName(i).Visible = (i < C)
                backPPLife(i).Visible = lblPPName(i).Visible
                backPPMana(i).Visible = lblPPName(i).Visible
                If lblPPName(i).Visible Then
                    lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                    shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                    shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
                    lblPPLife(i).Caption = "PV : " & Player(Val(lblPPName(i).Tag)).HP & "/" & Player(Val(lblPPName(i).Tag)).MaxHp
                    lblPPMana(i).Caption = "PM : " & Player(Val(lblPPName(i).Tag)).MP & "/" & Player(Val(lblPPName(i).Tag)).MaxMp
                End If
            Next
        End If
    Else: picParty.Visible = False: End If
End Sub

Private Sub Picture16_Click()
    If Player(MyIndex).PartyIndex > 0 Then
        Dim i As Integer, C As Byte
        C = 3
        For i = (Val(lblPPName(0).Tag) - 1) To 1 Step -1
            If i > 0 Then
                If IsPlaying(i) And Player(i).PartyIndex = Player(MyIndex).PartyIndex And C > 0 And i <> MyIndex Then
                    C = C - 1
                    lblPPName(C).Tag = i
                End If
            End If
        Next
        For i = 0 To 2
            lblPPName(i).Visible = (i <= Abs(C - 3))
            backPPLife(i).Visible = lblPPName(i).Visible
            backPPMana(i).Visible = lblPPName(i).Visible
            If lblPPName(i).Visible Then
                lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).level
                shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
                shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
                lblPPLife(i).Caption = "PV : " & Player(Val(lblPPName(i).Tag)).HP & "/" & Player(Val(lblPPName(i).Tag)).MaxHp
                lblPPMana(i).Caption = "PM : " & Player(Val(lblPPName(i).Tag)).MP & "/" & Player(Val(lblPPName(i).Tag)).MaxMp
            End If
            
            'lblPPName(i).Visible = True: backPPLife(i).Visible = True: backPPMana(i).Visible = True
            'lblPPName(i).Caption = Trim$(Player(Val(lblPPName(i).Tag)).name) & " - " & Player(Val(lblPPName(i).Tag)).Level
            'shpPPLife(i).Width = Player(Val(lblPPName(i).Tag)).HP / Player(Val(lblPPName(i).Tag)).MaxHp * backPPLife(i).Width
            'shpPPMana(i).Width = Player(Val(lblPPName(i).Tag)).MP / Player(Val(lblPPName(i).Tag)).MaxMp * backPPMana(i).Width
        Next
    Else: picParty.Visible = False: End If
End Sub

Private Sub Picture17_Click()
    Picture11.Top = Picture11.Top + 88
    
End Sub

Private Sub Picture18_Click()
    Picture11.Top = Picture11.Top - 88
    If Picture11.Top > 0 Then Picture11.Top = 0
End Sub

Private Sub Picture9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    itmDesc.Visible = False
End Sub

Private Sub qf_Click()
picquete.Visible = False
End Sub

Private Sub quetetimersec_Timer()
Dim Queten As Long

Queten = Val(Player(MyIndex).QueteEnCour)
If Queten <= 0 Then Exit Sub
If quete(Queten).Temps > 0 And Player(MyIndex).QueteEnCour > 0 Then

Seco = Seco - 1
If Seco <= 0 And Minu > 0 Then
    Seco = 59
    seconde.Caption = Seco
    Minu = Minu - 1
    If Len(STR$(Minu)) > 2 Then Minute.Caption = Minu & ":" Else Minute.Caption = "0" & Minu & ":"
End If
If Seco <= 0 And Minu <= 0 Then
    seconde.Caption = 0
    Call MsgBox("La quête : " & Trim$(quete(Queten).nom) & " est terminer, le temps est écoulé")
    Player(MyIndex).QueteEnCour = 0
    quetetimersec.Enabled = False
    tmpsquete.Visible = False
End If

If Len(STR$(Seco)) > 2 Then seconde.Caption = Seco Else seconde.Caption = "0" & Seco
lbltimeQuete.Visible = True
lbltimeQuete.Caption = "Quête se termine dans :" & Minu & " minute(s) et " & Seco & " seconde."
Else
Player(MyIndex).QueteEnCour = 0
tmpsquete.Visible = False
quetetimersec.Enabled = False
lbltimeQuete.Visible = False
End If

End Sub

Private Sub scrlBltText_Change()
Dim i As Long
    For i = 1 To MAX_BLT_LINE
        BattlePMsg(i).Index = 1
        BattlePMsg(i).Time = i
        BattleMMsg(i).Index = 1
        BattleMMsg(i).Time = i
    Next i
    
    MAX_BLT_LINE = scrlBltText.Value
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    lblLines.Caption = "Nbr de ligne écrite sur l'écran: " & scrlBltText.Value
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
    If IsConnected Then Call IncomingData(bytesTotal)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If ConOff = True Then Exit Sub
    Call HandleKeypresses(KeyAscii)
    If (KeyAscii = vbKeyReturn) Then KeyAscii = 0
    If (frmMirage.txtMyTextBox.Visible = False) And (KeyAscii = optTouche(CByte(Val(ReadINI("TJEU", "action", App.Path & "\Config\Option.ini")))).Value) Then KeyAscii = 0
    If KeyAscii = vbKeyEscape Then
        If fra_fenetre.Visible = True Then fra_fenetre.Visible = False
        If fra_info.Visible = True Then fra_info.Visible = False
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If ConOff = True Then Exit Sub
    Call CheckInput(1, KeyCode, Shift)
    On Error Resume Next
    txtMyTextBox.SetFocus
End Sub

Private Sub sync_Timer()
SendData ("sync" & END_CHAR)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Mediaplayer.URL > vbNullString Then
    If Mediaplayer.Controls.currentPosition = 0 And Mediaplayer.currentMedia.name = Mid$(Map(GetPlayerMap(MyIndex)).Music, 1, Len(Map(GetPlayerMap(MyIndex)).Music) - 4) Then Call frmMirage.Mediaplayer.Controls.Play
End If
End Sub

Private Sub Timer2_Timer()
    Call affrac
    'Timer2.Enabled = False
End Sub

Private Sub tmrRainDrop_Timer()
    If BLT_RAIN_DROPS > RainIntensity Then tmrRainDrop.Enabled = False: Exit Sub
    If BLT_RAIN_DROPS > 0 Then If DropRain(BLT_RAIN_DROPS).Randomized = False Then Call RNDRainDrop(BLT_RAIN_DROPS)
    BLT_RAIN_DROPS = BLT_RAIN_DROPS + 1
    If tmrRainDrop.Interval > 30 Then tmrRainDrop.Interval = tmrRainDrop.Interval - 10
End Sub

Private Sub tmrSnowDrop_Timer()
    If BLT_SNOW_DROPS > RainIntensity Then tmrSnowDrop.Enabled = False: Exit Sub
    If BLT_SNOW_DROPS > 0 Then If DropSnow(BLT_SNOW_DROPS).Randomized = False Then Call RNDSnowDrop(BLT_SNOW_DROPS)
    BLT_SNOW_DROPS = BLT_SNOW_DROPS + 1
    If tmrSnowDrop.Interval > 30 Then tmrSnowDrop.Interval = tmrSnowDrop.Interval - 10
End Sub

Private Sub lblUseItem_Click()
Dim d As Long

If Player(MyIndex).Inv(Inventory).num <= 0 Then Exit Sub

Call SendUseItem(Inventory)

For d = 1 To MAX_INV
    If Player(MyIndex).Inv(d).num > 0 Then If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then picInv(d - 1).Picture = LoadPicture()
Next d

Call UpdateVisInv
End Sub

Private Sub lblDropItem_Click()
    Call DropItems
End Sub

Sub DropItems()
Dim InvNum As Long
Dim GoldAmount As String
On Error GoTo Done

If Inventory <= 0 Then Exit Sub
InvNum = Inventory
   
    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Empilable <> 0 Then
            GoldAmount = InputBox("Combien de " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") voulez vous jeter?", "Jeter " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmMirage.Left, frmMirage.Top)
            If IsNumeric(GoldAmount) Then Call SendDropItem(InvNum, GoldAmount)
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
   
    picInv(InvNum - 1).Picture = LoadPicture()
    Call UpdateVisInv
    Exit Sub
Done:
    If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then MsgBox "Trop grande quantiter(erreur du logiciel)"
End Sub

Private Sub cmdAccess_Click()
Dim Packet As String
    If txtName.Text = vbNullString Or txtAccess.Text = vbNullString Or Not IsNumeric(txtAccess.Text) Then Exit Sub
    Packet = "GUILDCHANGEACCESS" & SEP_CHAR & txtName.Text & SEP_CHAR & txtAccess.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdDisown_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "GUILDDISOWN" & SEP_CHAR & txtName.Text & END_CHAR
    Call SendData(Packet)
End Sub

Private Sub cmdTrainee_Click()
Dim Packet As String
    If txtName.Text = vbNullString Then Exit Sub
    Packet = "guildtraineevbyesno" & SEP_CHAR & txtName.Text & END_CHAR '"GUILDTRAINEE"
    Call SendData(Packet)
End Sub

Private Sub txtQ_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub txtQ_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 4
DragX = x
DragY = y
End Sub

Private Sub txtQ_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 4 Then txtQ.Top = txtQ.Top + ((y / twippy) - (DragY / twippy)): txtQ.Left = txtQ.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub txtQ_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub TxtQ2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: txtQ.Visible = False
End Sub

Private Sub txtTempsBulles_Change()
If IsNumeric(txtTempsBulles.Text) Then
WriteINI "CONFIG", "bubbletime", txtTempsBulles, App.Path & "\Config\Client.ini"
End If
End Sub

Private Sub Up_Click()
If VScroll1.Value = 0 Then Exit Sub
    VScroll1.Value = VScroll1.Value - 1
    Picture9.Top = Picture9.Top + 88 'VScroll1.value * -PIC_Y
End Sub

Private Sub Down_Click()
Dim x As Byte
x = Int(MAX_INV / 8)
If x * 8 < MAX_INV Then x = x + 1
If VScroll1.Value = x Then Exit Sub
    VScroll1.Value = VScroll1.Value + 1
    Picture9.Top = Picture9.Top - 88 'VScroll1.value * -PIC_Y
End Sub

Public Sub ClearPic()
    fra_info.Visible = False
    picquete.Visible = False
    picEquip.Visible = False
    picInv3.Visible = False
    picPlayerSpells.Visible = False
    picWhosOnline.Visible = False
    picGuild.Visible = False
    picGuildAdmin.Visible = False
    picGuild.Visible = False
End Sub

