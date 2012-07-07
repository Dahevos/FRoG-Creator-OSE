VERSION 5.00
Begin VB.Form frmNpcEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer un PNJ (PNJ = personnage non joueur)"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   90
   ClientWidth     =   11535
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
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   769
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSpells 
      Caption         =   "Sortillèges"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5400
      TabIndex        =   57
      Top             =   5520
      Width           =   6015
      Begin VB.ComboBox cmbSpellType 
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
         ItemData        =   "frmNpcEditor.frx":0000
         Left            =   120
         List            =   "frmNpcEditor.frx":0025
         TabIndex        =   62
         Text            =   "Type de Sort"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ListBox lstTypeSpell 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   2655
      End
      Begin VB.ListBox lstSpells 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         IntegralHeight  =   0   'False
         Left            =   3240
         TabIndex        =   60
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdAddSpell 
         Caption         =   ">"
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
         Left            =   2880
         TabIndex        =   59
         Top             =   720
         Width           =   255
      End
      Begin VB.CommandButton cmdRemSpell 
         Caption         =   "<"
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
         Left            =   2880
         TabIndex        =   58
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Divers"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   5400
      TabIndex        =   31
      Top             =   1200
      Width           =   6015
      Begin VB.HScrollBar quetenum 
         Height          =   255
         LargeChange     =   10
         Left            =   1560
         Max             =   10000
         Min             =   1
         TabIndex        =   52
         Top             =   3600
         Value           =   1
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtSpawnSecs 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   50
         Text            =   "0"
         ToolTipText     =   "Temps mit par le PNJ pour ressusciter après sa mort"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmbBehavior 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmNpcEditor.frx":00CC
         Left            =   840
         List            =   "frmNpcEditor.frx":00E5
         Style           =   2  'Dropdown List
         TabIndex        =   48
         ToolTipText     =   "Aptitude que doit avoir le PNJ face aux joueurs"
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CheckBox chkDay 
         Caption         =   "Jour"
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
         Left            =   5160
         TabIndex        =   45
         ToolTipText     =   "Si cette case est cochée le PNJ apparaîtras le jour"
         Top             =   2760
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkNight 
         Caption         =   "Nuit"
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
         Left            =   5160
         TabIndex        =   46
         ToolTipText     =   "Si cette case est cochée le PNJ apparaîtras la nuit"
         Top             =   2520
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.TextBox txtChance 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         TabIndex        =   35
         Text            =   "0"
         ToolTipText     =   "Chance pour le joueur d'avoir l'objet quand il tue le PNJ"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   1080
         Max             =   10000
         TabIndex        =   34
         Top             =   1680
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   1080
         Max             =   500
         TabIndex        =   33
         Top             =   1200
         Value           =   1
         Width           =   3255
      End
      Begin VB.HScrollBar scrlDropItem 
         Height          =   255
         Left            =   1080
         Max             =   5
         Min             =   1
         TabIndex        =   32
         Top             =   240
         Value           =   1
         Width           =   3255
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Numéro de la quête :"
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
         TabIndex        =   54
         ToolTipText     =   "Nombre d'objet donné"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label qutn 
         AutoSize        =   -1  'True
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
         Left            =   4320
         TabIndex        =   53
         ToolTipText     =   "Nombre d'objet donné"
         Top             =   3600
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Apparition :"
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
         Left            =   4800
         TabIndex        =   51
         ToolTipText     =   "Moment ou le PNJ apparaîtras sur la carte"
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Temps entre chaque réapparition :"
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
         TabIndex        =   49
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Atitude :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   47
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "1 chance de l'avoir sur..."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   44
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
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
         Left            =   4440
         TabIndex        =   43
         ToolTipText     =   "Nombre d'objet donné"
         Top             =   1680
         Width           =   75
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Valeur :"
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
         TabIndex        =   42
         ToolTipText     =   "Nombre d'objet donné"
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
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
         Left            =   4440
         TabIndex        =   41
         ToolTipText     =   "Numéros de l'objet donné"
         Top             =   1200
         Width           =   75
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Numéro :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   "Numéros de l'objet donné"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblItemName 
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
         Left            =   960
         TabIndex        =   39
         ToolTipText     =   "Nom de l'objet donné"
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Objet :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Nom de l'objet donné"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDropItem 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   4440
         TabIndex        =   37
         ToolTipText     =   "Numéros de l'objet donné par le PNJ a sa mort : un PNJ peut donner 10 objet différent au maximum"
         Top             =   240
         Width           =   75
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Objet donné :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         ToolTipText     =   "Numéros de l'objet donné par le PNJ a sa mort : un PNJ peut donner 10 objet différent au maximum"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informations Générales"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   1080
         ScaleHeight     =   83
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   93
         TabIndex        =   65
         Top             =   660
         Width           =   1425
         Begin VB.PictureBox picSprite 
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
            Height          =   480
            Left            =   0
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   66
            ToolTipText     =   "Sprite du PNJ"
            Top             =   0
            Width           =   480
         End
      End
      Begin VB.VScrollBar scrlSpriteY 
         Height          =   1275
         Left            =   2580
         Max             =   1
         TabIndex        =   64
         Top             =   660
         Width           =   255
      End
      Begin VB.HScrollBar scrlSpriteX 
         Height          =   255
         Left            =   1080
         Max             =   1
         TabIndex        =   63
         Top             =   1980
         Width           =   1455
      End
      Begin VB.HScrollBar StartHP 
         Height          =   255
         LargeChange     =   100
         Left            =   1080
         TabIndex        =   26
         Top             =   5400
         Width           =   2895
      End
      Begin VB.CheckBox vol 
         Caption         =   "PNJ volant"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3060
         TabIndex        =   56
         ToolTipText     =   "Si cette case est cochée les PNJ pouront passer à travers toutes les cases bloquer sauf celle pour les PNJs"
         Top             =   660
         Width           =   1095
      End
      Begin VB.CheckBox invi 
         Caption         =   "Invincible"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1080
         TabIndex        =   55
         ToolTipText     =   "Si cette case est cochée le PNJ ne pourat pas être tuer"
         Top             =   5640
         Width           =   1095
      End
      Begin VB.HScrollBar ExpGive 
         Height          =   255
         LargeChange     =   100
         Left            =   1080
         TabIndex        =   30
         Top             =   6000
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   1000
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         LargeChange     =   5
         Left            =   1080
         Max             =   30
         TabIndex        =   11
         Top             =   2400
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSTR 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   5000
         TabIndex        =   10
         Top             =   3000
         Width           =   2895
      End
      Begin VB.HScrollBar scrlDEF 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   5000
         TabIndex        =   9
         Top             =   3600
         Width           =   2895
      End
      Begin VB.HScrollBar scrlSPEED 
         Enabled         =   0   'False
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   5000
         TabIndex        =   8
         Top             =   4200
         Width           =   2895
      End
      Begin VB.HScrollBar scrlMAGI 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   5000
         TabIndex        =   7
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label lblExpGiven 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         ToolTipText     =   "Expérience donnée aux joueurs quand ils tuent le PNJ"
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "EXP donnée :"
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
         TabIndex        =   28
         ToolTipText     =   "Expérience donnée aux joueurs quand ils tuent le PNJ"
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label lblStartHP 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         ToolTipText     =   "Point de vie du PNJ"
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Points de Vie :"
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
         Left            =   50
         TabIndex        =   25
         ToolTipText     =   "Point de vie du PNJ"
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label lblSprite 
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
         Height          =   375
         Left            =   4080
         TabIndex        =   24
         ToolTipText     =   "Numéros du sprinte du PNJ (sprinte = habit/skin)"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Apparence :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblRange 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   22
         ToolTipText     =   "Distance en case ou le PNJ va aider les autres PNJ a ce défendre"
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Distance de vision:"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "Distance en case ou le PNJ va aider les autres PNJ a ce défendre"
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblSTR 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         ToolTipText     =   "Force du PNJ"
         Top             =   3000
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Force :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   19
         ToolTipText     =   "Force du PNJ"
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label lblDEF 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   18
         ToolTipText     =   "Défense du PNJ"
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Défense :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   17
         ToolTipText     =   "Défense du PNJ"
         Top             =   3600
         Width           =   615
      End
      Begin VB.Label lblSPEED 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   16
         ToolTipText     =   "Vitesse du PNJ"
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Vitesse :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         ToolTipText     =   "Vitesse du PNJ"
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label lblMAGI 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         ToolTipText     =   "Magie du PNJ"
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Magie :"
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
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Magie du PNJ"
         Top             =   4800
         Width           =   615
      End
   End
   Begin VB.TextBox txtAttackSay 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      ToolTipText     =   "Message dit par le pnj quand un joueur l'interpelle"
      Top             =   240
      Width           =   4335
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
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer le PNJ"
      Top             =   7920
      Width           =   1695
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
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Quitte la fenêtre d'édition et enregistre le PNJ"
      Top             =   7920
      Width           =   1695
   End
   Begin VB.TextBox txtName 
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
      Left            =   960
      TabIndex        =   0
      ToolTipText     =   "Nom du PNJ"
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Discussion :"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmNpcEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDay_Click()
    If chkNight.value = Unchecked Then If chkDay.value = Unchecked Then chkDay.value = Checked
End Sub

Private Sub chkNight_Click()
    If chkDay.value = Unchecked Then If chkNight.value = Unchecked Then chkNight.value = Checked
End Sub

Private Sub cmbBehavior_Click()
If cmbBehavior.ListIndex = 5 Then
    qutn.Visible = True
    Label20.Visible = True
    quetenum.Visible = True
    Label20.Caption = "Numéro de quête:"
    quetenum.Min = 1
    quetenum.Max = MAX_QUETES
ElseIf cmbBehavior.ListIndex = 6 Then
    qutn.Visible = True
    Label20.Visible = True
    quetenum.Visible = True
    Label20.Caption = "Case Script:"
    quetenum.Min = 1
    quetenum.Max = 255
ElseIf cmbBehavior.ListIndex = 3 Then
    qutn.Visible = True
    Label20.Visible = True
    quetenum.Visible = True
    Label20.Caption = "Magasin:"
    quetenum.Min = 1
    quetenum.Max = MAX_SHOPS
ElseIf cmbBehavior.ListIndex = 0 Or cmbBehavior.ListIndex = 1 Then
    qutn.Visible = True
    Label20.Visible = True
    quetenum.Visible = True
    Label20.Caption = "Type d'Arme:"
    quetenum.Min = 0
    quetenum.Max = 11 + MAX_METIER
Else
    qutn.Visible = False
    Label20.Visible = False
    quetenum.Visible = False
End If
End Sub

Private Sub cmbSpellType_Click()
Dim i As Integer
lstTypeSpell.Clear
For i = 1 To MAX_SPELLS
    If Spell(i).Type = cmbSpellType.ListIndex And (Trim$(Spell(i).name) <> vbNullString And Trim$(Spell(i).name) <> Space$(1)) Then lstTypeSpell.AddItem lstTypeSpell.ListCount + 1 & ". " & Trim$(Spell(i).name): lstTypeSpell.ItemData(lstTypeSpell.ListCount - 1) = i
Next
Debug.Print Trim$(Spell(6).name)
If lstTypeSpell.ListCount = 0 Then lstTypeSpell.AddItem "<Aucun>"
End Sub

Private Sub cmdAddSpell_Click()
If lstSpells.ListCount < MAX_NPC_SPELLS Then
If lstTypeSpell.ListIndex >= 0 Then
If lstTypeSpell.ItemData(lstTypeSpell.ListIndex) > 0 Then If Not InItemData(lstSpells, lstTypeSpell.ItemData(lstTypeSpell.ListIndex)) Then lstSpells.AddItem lstSpells.ListCount + 1 & ". " & Spell(lstTypeSpell.ItemData(lstTypeSpell.ListIndex)).name: lstSpells.ItemData(lstSpells.ListCount - 1) = lstTypeSpell.ItemData(lstTypeSpell.ListIndex)
End If
Else
MsgBox "Il est impossible d'ajouter plus de sorts ." & vbCrLf & "Maximum : " & MAX_NPC_SPELLS
End If
End Sub

Private Function InItemData(List As ListBox, ItemDataValue As Integer) As Boolean
Dim i As Integer
For i = 0 To List.ListCount - 1
    If List.ItemData(i) = ItemDataValue Then InItemData = True: Exit Function
Next
InItemData = False
End Function

Private Sub cmdRemSpell_Click()
If lstSpells.ListIndex >= 0 Then
lstSpells.ItemData(lstSpells.ListIndex) = 0
lstSpells.RemoveItem lstSpells.ListIndex
End If
End Sub

Private Sub ExpGive_Change()
    lblExpGiven.Caption = ExpGive.value
End Sub

Private Sub ExpGive_Scroll()
    ExpGive_Change
End Sub

Private Sub Form_Load()
    scrlDropItem.Max = MAX_NPC_DROPS
    'picSprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png")
    quetenum.Max = MAX_QUETES
    'picSprite.Height = (48 * Screen.TwipsPerPixelY)
    'picSprite.Top = Picture1.Top + 30
    'picSprite.Left = Picture1.Left + 30
    'Picture1.Height = (48 * Screen.TwipsPerPixelY) + 44
    'Picture1.Width = (32 * Screen.TwipsPerPixelY) + 44
    scrlSpriteY.Max = picSprite.Height - Picture1.Height
    scrlSpriteX.Max = picSprite.Width - Picture1.Width
    Call AffSprites
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call NpcEditorCancel
End Sub

Private Sub invi_Click()
If StartHP.Enabled Then StartHP.Enabled = False: StartHP.value = 0 Else StartHP.Enabled = True: StartHP.value = 1
End Sub

Private Sub quetenum_Change()
Npc(EditorIndex).quetenum = quetenum.value
If cmbBehavior.ListIndex = 0 Or cmbBehavior.ListIndex = 1 Then
    If quetenum.value = 0 Then
         qutn.Caption = "Toute Arme"
    ElseIf quetenum.value = 1 Then
         qutn.Caption = "Epées"
    ElseIf quetenum.value = 2 Then
         qutn.Caption = "Haches"
    ElseIf quetenum.value = 3 Then
         qutn.Caption = "Dagues"
    ElseIf quetenum.value = 4 Then
         qutn.Caption = "Faux"
    ElseIf quetenum.value = 5 Then
         qutn.Caption = "Marteaux"
    ElseIf quetenum.value = 6 Then
         qutn.Caption = "Pioches"
    ElseIf quetenum.value = 7 Then
         qutn.Caption = "Pelles"
    ElseIf quetenum.value = 8 Then
         qutn.Caption = "Batons"
    ElseIf quetenum.value = 9 Then
         qutn.Caption = "Baguettes"
    ElseIf quetenum.value = 10 Then
         qutn.Caption = "Outillages"
    ElseIf quetenum.value = 11 Then
         qutn.Caption = "Arc"
    ElseIf quetenum.value > 11 Then
        qutn.Caption = "Metier: " & Metier(quetenum.value - 11).nom
    End If
Else
    qutn.Caption = quetenum.value
End If
End Sub

Private Sub scrlDEF_Scroll()
    scrlDEF_Change
End Sub

Private Sub scrlDropItem_Change()
    txtChance.Text = Npc(EditorIndex).ItemNPC(scrlDropItem.value).Chance
    scrlNum.value = Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemNum
    scrlValue.value = Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemValue
    lblDropItem.Caption = scrlDropItem.value
End Sub

Private Sub scrlMAGI_Scroll()
    scrlMAGI_Change
End Sub

Private Sub scrlRange_Scroll()
    scrlRange_Change
End Sub

Private Sub scrlSPEED_Scroll()
    scrlSPEED_Change
End Sub

Private Sub scrlSprite_Change()
    lblSprite.Caption = CStr(scrlSprite.value)
    Call AffSprites
End Sub

Private Sub scrlRange_Change()
    lblRange.Caption = CStr(scrlRange.value)
End Sub

Private Sub scrlSprite_Scroll()
    scrlSprite_Change
End Sub

Private Sub scrlSpriteX_Change()
    picSprite.Left = scrlSpriteX.value
End Sub

Private Sub scrlSpriteY_Change()
    picSprite.Top = scrlSpriteY.value
End Sub

Private Sub scrlSTR_Change()
    lblSTR.Caption = CStr(scrlSTR.value)
End Sub

Private Sub scrlDEF_Change()
    lblDEF.Caption = CStr(scrlDEF.value)
End Sub

Private Sub scrlSPEED_Change()
    lblSPEED.Caption = CStr(scrlSPEED.value)
End Sub

Private Sub scrlMAGI_Change()
    lblMAGI.Caption = CStr(scrlMAGI.value)
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = CStr(scrlNum.value)
    lblItemName.Caption = vbNullString
    If scrlNum.value > 0 Then lblItemName.Caption = Trim$(Item(scrlNum.value).name): lblItemName.ForeColor = Item(scrlNum.value).NCoul
    Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemNum = scrlNum.value
End Sub

Private Sub scrlSTR_Scroll()
    scrlSTR_Change
End Sub

Private Sub scrlValue_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.value).ItemValue = scrlValue.value
    lblValue.Caption = CStr(scrlValue.value)
End Sub

Private Sub cmdOk_Click()
Dim tmp As Integer
    If StartHP.value <= 0 And Not CBool(invi.value) Then
        tmp = MsgBox("ATTENTION : Le PNJ n'a pas de points de vie." & vbCrLf & "Il sera donc considérer comme mort, et ne pourrat donc pas parler. Voulez-vous continuer?", vbYesNo, "ATTENTION")
        If tmp = vbNo Then Exit Sub
    End If
    Call NpcEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call NpcEditorCancel
End Sub

Private Sub StartHP_Change()
    lblStartHP.Caption = StartHP.value
End Sub

Private Sub StartHP_Scroll()
    StartHP_Change
End Sub

Private Sub txtChance_Change()
    Npc(EditorIndex).ItemNPC(scrlDropItem.value).Chance = Val(txtChance.Text)
End Sub

Private Sub AffSprites()
On Error Resume Next
Call PrepareSprite(scrlSprite.value)
picSprite.Height = (DDSD_Character(scrlSprite.value).lHeight) / 4
picSprite.Width = (DDSD_Character(scrlSprite.value).lWidth) / 4
scrlSpriteY.Max = (picSprite.Height) - Picture1.Height
scrlSpriteX.Max = (picSprite.Width) - Picture1.Width
If picSprite.Height > 96 Then
scrlSpriteY.Visible = True
Else
scrlSpriteY.Visible = False
End If
If picSprite.Width > 96 Then
scrlSpriteX.Visible = True
Else
scrlSpriteX.Visible = False
End If
'Picture1.Height = picSprite.Height + 60
'Picture1.Width = picSprite.Width + 60
Call AffSurfPic(DD_SpriteSurf(scrlSprite.value), picSprite, 0, 0)
scrlSprite.Max = MAX_DX_SPRITE
End Sub
