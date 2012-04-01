VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmItemEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditeur d'objets"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   285
   ClientWidth     =   11355
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
   ScaleHeight     =   447
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   757
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSelect 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2820
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   4
      ToolTipText     =   "Image qui sera affiché dans l'inventaire "
      Top             =   2820
      Width           =   465
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   11827
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1984
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Objet"
      TabPicture(0)   =   "frmItemEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraPet"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraobjsc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraSpell"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraVitals"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "framonture"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "VScroll1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtDesc"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fraBow"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "coulr"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "coul"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FramePD"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "picPic"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtName"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmbType"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdOk"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "FraOption"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraAttributes"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "fraEquipment"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CtArme"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdCancel"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6780
         TabIndex        =   99
         ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer l'objet"
         Top             =   6120
         Width           =   1155
      End
      Begin VB.ComboBox CtArme 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmItemEditor.frx":001C
         Left            =   8040
         List            =   "frmItemEditor.frx":0044
         Style           =   2  'Dropdown List
         TabIndex        =   112
         ToolTipText     =   "Sélectionner un type pour l'objet"
         Top             =   540
         Width           =   3195
      End
      Begin VB.Frame fraEquipment 
         Caption         =   "Caractéristiques de l'objet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   5580
         TabIndex        =   64
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.HScrollBar scrlSexReq 
            Height          =   255
            Left            =   300
            Max             =   2
            TabIndex        =   111
            Top             =   4080
            Value           =   2
            Width           =   1875
         End
         Begin VB.HScrollBar scrlDurability 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   10000
            Min             =   -1
            TabIndex        =   71
            Top             =   480
            Value           =   -1
            Width           =   1875
         End
         Begin VB.HScrollBar scrlStrength 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   10000
            TabIndex        =   70
            Top             =   1080
            Value           =   1
            Width           =   1875
         End
         Begin VB.HScrollBar scrlStrReq 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   10000
            TabIndex        =   69
            Top             =   1680
            Width           =   1875
         End
         Begin VB.HScrollBar scrlDefReq 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   10000
            TabIndex        =   68
            Top             =   2280
            Width           =   1875
         End
         Begin VB.HScrollBar scrlSpeedReq 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   10000
            TabIndex        =   67
            Top             =   2880
            Width           =   1875
         End
         Begin VB.HScrollBar scrlClassReq 
            Height          =   255
            Left            =   300
            Max             =   1
            Min             =   -1
            TabIndex        =   66
            Top             =   3480
            Value           =   -1
            Width           =   1875
         End
         Begin VB.HScrollBar scrlAccessReq 
            Height          =   255
            Left            =   300
            Max             =   4
            TabIndex        =   65
            Top             =   4620
            Width           =   1875
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Les 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1890
            TabIndex        =   110
            ToolTipText     =   "Classe requise"
            Top             =   3840
            Width           =   315
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sex Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   109
            ToolTipText     =   "Classe requise"
            Top             =   3840
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Durabilité :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   85
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Dommage :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   84
            ToolTipText     =   "Dommage infliger par l'objet"
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblDurability 
            Alignment       =   1  'Right Justify
            Caption         =   "Ind."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1500
            TabIndex        =   83
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblStrength 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   82
            ToolTipText     =   "Dommage infliger par l'objet"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Force Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   300
            TabIndex        =   81
            ToolTipText     =   "Force requise"
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Defense Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   80
            ToolTipText     =   "Défense requise"
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   79
            ToolTipText     =   "Force requise"
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   78
            ToolTipText     =   "Défense requise"
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   77
            ToolTipText     =   "Vitesse requise"
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vitesse Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   76
            ToolTipText     =   "Vitesse requise"
            Top             =   2640
            Width           =   975
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Classe Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   75
            ToolTipText     =   "Classe requise"
            Top             =   3240
            Width           =   975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Access Req :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   74
            ToolTipText     =   "Accès requit"
            Top             =   4380
            Width           =   975
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aucune"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1080
            TabIndex        =   73
            ToolTipText     =   "Classe requise"
            Top             =   3240
            Width           =   1125
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tous"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   180
            TabIndex        =   72
            ToolTipText     =   "Accès requit"
            Top             =   4380
            Width           =   1935
         End
      End
      Begin VB.Frame fraAttributes 
         Caption         =   "Attribut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5715
         Left            =   8040
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   3195
         Begin VB.HScrollBar scrlAttackSpeed 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   5000
            Min             =   1
            TabIndex        =   40
            Top             =   5340
            Value           =   1000
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddEXP 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   100
            TabIndex        =   27
            Top             =   2460
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSP 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   25
            Top             =   1800
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddSpeed 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   17
            Top             =   4800
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMagi 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   16
            Top             =   4200
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddDef 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   15
            Top             =   3600
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddStr 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   14
            Top             =   3000
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   13
            Top             =   1200
            Width           =   2655
         End
         Begin VB.HScrollBar scrlAddHP 
            Height          =   230
            LargeChange     =   10
            Left            =   360
            Max             =   10000
            Min             =   -10000
            TabIndex        =   12
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label lblAttackSpeed 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1000 Milleseconds"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1500
            TabIndex        =   42
            ToolTipText     =   "Intervalle entre 2 coups d'une arme en milliseconde (1000milliseconde = 1seconde)"
            Top             =   5100
            Width           =   1110
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vitesse d'attaque :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            ToolTipText     =   "Intervalle entre 2 coups d'une arme en milliseconde (1000milliseconde = 1seconde)"
            Top             =   5100
            Width           =   1215
         End
         Begin VB.Label lblAddEXP 
            Alignment       =   1  'Right Justify
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   29
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ajout d'EXP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   2220
            Width           =   855
         End
         Begin VB.Label lblAddSP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   26
            Top             =   1560
            Width           =   495
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de SP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label lblAddSpeed 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   23
            ToolTipText     =   "Ajouter de la vitesse"
            Top             =   4560
            Width           =   495
         End
         Begin VB.Label lblAddMagi 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   22
            ToolTipText     =   "Ajouter de la magie"
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label lblAddDef 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   21
            ToolTipText     =   "Ajouter de la défense "
            Top             =   3360
            Width           =   495
         End
         Begin VB.Label lblAddStr 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   20
            ToolTipText     =   "Ajouter de la force"
            Top             =   2760
            Width           =   495
         End
         Begin VB.Label lblAddMP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   19
            ToolTipText     =   "Ajouter des points de magie"
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblAddHP 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   18
            ToolTipText     =   "Ajouter des points de vie"
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ajout de Vit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Ajouter de la vitesse"
            Top             =   4560
            Width           =   855
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de Mag"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            ToolTipText     =   "Ajouter de la magie"
            Top             =   3960
            Width           =   855
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de Def"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Ajouter de la défense "
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de For"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Ajouter de la force"
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de MP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Ajouter des points de magie"
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Ajout de HP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Ajouter des points de vie"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame FraOption 
         Caption         =   "Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   101
         Top             =   2400
         Width           =   2115
         Begin VB.CheckBox CheckEmpi 
            Caption         =   "Objet Empilable"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   1875
         End
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   100
         ToolTipText     =   "Quitte la fenêtre d'édition et enregistre l'objet"
         Top             =   6120
         Width           =   1155
      End
      Begin VB.ComboBox cmbType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmItemEditor.frx":00AE
         Left            =   5580
         List            =   "frmItemEditor.frx":00E5
         Style           =   2  'Dropdown List
         TabIndex        =   98
         ToolTipText     =   "Sélectionner un type pour l'objet"
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   97
         ToolTipText     =   "Nom de l'objet"
         Top             =   540
         Width           =   5175
      End
      Begin VB.PictureBox picPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Left            =   180
         ScaleHeight     =   215
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   192
         TabIndex        =   96
         ToolTipText     =   "Sélectionner une image pour l'objet"
         Top             =   3360
         Width           =   2880
      End
      Begin VB.Frame FramePD 
         Caption         =   "Paperdoll"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   3360
         TabIndex        =   59
         Top             =   3060
         Visible         =   0   'False
         Width           =   2115
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1020
            Left            =   540
            ScaleHeight     =   66
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   66
            TabIndex        =   62
            Top             =   900
            Width           =   1020
            Begin VB.PictureBox PicPD 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   15
               ScaleHeight     =   64
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   64
               TabIndex        =   63
               ToolTipText     =   "Image qui sera affiché dans l'inventaire "
               Top             =   15
               Width           =   960
            End
         End
         Begin VB.HScrollBar scrlPD 
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   540
            Width           =   1575
         End
         Begin VB.CheckBox CheckPD 
            Caption         =   "Paperdoll"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblpaper 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   108
            Top             =   540
            Width           =   315
         End
      End
      Begin VB.PictureBox coul 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   55
         ToolTipText     =   "Couleur de l'objet"
         Top             =   2340
         Width           =   375
      End
      Begin MSComDlg.CommonDialog cmd 
         Left            =   3480
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Couleur de l'objet"
      End
      Begin VB.CommandButton coulr 
         Caption         =   "Définir la couleur de l'objet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   54
         ToolTipText     =   "La couleur peut servir à définir la rareté de l'objet ou sa puissance le nom de l'objet sera de cette couleur"
         Top             =   2340
         Width           =   2595
      End
      Begin VB.Frame fraBow 
         Caption         =   "Arc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3360
         TabIndex        =   32
         Top             =   5040
         Visible         =   0   'False
         Width           =   2115
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   120
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   34
            TabIndex        =   35
            Top             =   840
            Width           =   540
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00404040&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   15
               ScaleHeight     =   32
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   32
               TabIndex        =   36
               Top             =   15
               Width           =   480
               Begin VB.PictureBox picBow 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00000000&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   480
                  Left            =   -1020
                  ScaleHeight     =   32
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   128
                  TabIndex        =   37
                  ToolTipText     =   "Image de la flèche"
                  Top             =   0
                  Width           =   1920
               End
            End
         End
         Begin VB.ComboBox cmbBow 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmItemEditor.frx":01C8
            Left            =   120
            List            =   "frmItemEditor.frx":01CA
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Sélectionner une flèche"
            Top             =   480
            Width           =   1935
         End
         Begin VB.CheckBox chkBow 
            Caption         =   "Arc"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Cocher la case si l'arme et un arc"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblName 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   660
            TabIndex        =   39
            ToolTipText     =   "Nom de la flèche"
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nom :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   720
            TabIndex        =   38
            Top             =   840
            Width           =   390
         End
      End
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         ToolTipText     =   "Description de l'objet"
         Top             =   1140
         Width           =   5175
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3240
         LargeChange     =   10
         Left            =   3060
         Max             =   464
         TabIndex        =   3
         Top             =   3360
         Width           =   255
      End
      Begin VB.Frame framonture 
         Caption         =   "Monture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   5580
         TabIndex        =   43
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.HScrollBar vit 
            Height          =   255
            Left            =   360
            Max             =   8
            Min             =   1
            TabIndex        =   56
            Top             =   2100
            Value           =   1
            Width           =   1875
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1020
            Left            =   720
            ScaleHeight     =   66
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   66
            TabIndex        =   47
            Top             =   720
            Width           =   1020
            Begin VB.PictureBox imgmont 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   960
               Left            =   15
               ScaleHeight     =   64
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   64
               TabIndex        =   48
               ToolTipText     =   "Image de la flèche"
               Top             =   15
               Width           =   960
            End
         End
         Begin VB.HScrollBar skin 
            Height          =   255
            LargeChange     =   10
            Left            =   360
            Max             =   1000
            Min             =   1
            TabIndex        =   44
            Top             =   420
            Value           =   1
            Width           =   1875
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Multiplie la vitesse par :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   1860
            Width           =   1680
         End
         Begin VB.Label nbvit 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1740
            TabIndex        =   57
            Top             =   1860
            Width           =   510
         End
         Begin VB.Label numskin 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1860
            TabIndex        =   46
            Top             =   180
            Width           =   390
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Sprite :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   180
            Width           =   525
         End
      End
      Begin VB.Frame fraVitals 
         Caption         =   "Modification des points apportés"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5580
         TabIndex        =   86
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.HScrollBar scrlVitalMod 
            Height          =   255
            LargeChange     =   10
            Left            =   300
            Max             =   1000
            TabIndex        =   87
            Top             =   660
            Value           =   1
            Width           =   1875
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Modifie les poitns de :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   89
            ToolTipText     =   "Modifie les points spécifier par le type de l'objet"
            Top             =   420
            Width           =   1395
         End
         Begin VB.Label lblVitalMod 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1560
            TabIndex        =   88
            ToolTipText     =   "Modifie les points spécifier par le type de l'objet"
            Top             =   420
            Width           =   615
         End
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Caractéristiques du sort"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   5580
         TabIndex        =   90
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   255
            Min             =   1
            TabIndex        =   91
            Top             =   1140
            Value           =   1
            Width           =   1875
         End
         Begin VB.Label lblSpellName 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   540
            Width           =   2160
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Numéro du sort :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   94
            ToolTipText     =   "Numéros du sort"
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label lblSpell 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   900
            TabIndex        =   95
            ToolTipText     =   "Numéros du sort"
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Nom du sort :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   93
            ToolTipText     =   "Nom du sort"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame fraobjsc 
         Caption         =   "Objet Scripter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   5580
         TabIndex        =   49
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.CheckBox disp 
            Caption         =   "L'objet disparaît de l'inventaire"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   240
            TabIndex        =   53
            ToolTipText     =   "Détermine si l'objet disparaîtra de l'inventaire quand le joueur l'utilisera"
            Top             =   1200
            Width           =   1875
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            TabIndex        =   50
            Top             =   720
            Width           =   1875
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1380
            TabIndex        =   51
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Label31 
            Caption         =   "Case de script :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   240
            TabIndex        =   52
            ToolTipText     =   "Numéros de la case qui se déclencheras quand le joueur équipera l'objet"
            Top             =   360
            Width           =   1965
         End
      End
      Begin VB.Frame fraPet 
         Caption         =   "Famillier"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   5580
         TabIndex        =   103
         Top             =   900
         Visible         =   0   'False
         Width           =   2355
         Begin VB.HScrollBar scrlPet 
            Height          =   255
            Left            =   180
            Min             =   1
            TabIndex        =   106
            Top             =   900
            Value           =   1
            Width           =   1695
         End
         Begin VB.Label lblPetNum 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1860
            TabIndex        =   107
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lblPetNom 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   105
            Top             =   540
            Width           =   1995
         End
         Begin VB.Label Label32 
            Caption         =   "Nom :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   104
            Top             =   300
            Width           =   1035
         End
      End
      Begin VB.Label Label26 
         Caption         =   "Description :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Nom de l'objet :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Image de l'objet :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   2940
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmItemEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VitVal As Byte

Private Sub chkBow_Click()
Dim i As Long
    If chkBow.value = Unchecked Then
        cmbBow.Clear
        cmbBow.AddItem "Aucune", 0
        cmbBow.ListIndex = 0
        cmbBow.Enabled = False
        lblName.Caption = vbNullString
    Else
        cmbBow.Clear
        For i = 1 To MAX_ARROWS
            cmbBow.AddItem i & " : " & Arrows(i).name
        Next i
        cmbBow.ListIndex = 0
        cmbBow.Enabled = True
    End If
End Sub

Private Sub cmbBow_Click()
    lblName.Caption = Arrows(cmbBow.ListIndex + 1).name
    Call AffSurfPic(DD_ArrowAnim, picBow, PIC_X, Arrows(cmbBow.ListIndex + 1).Pic * PIC_Y)
    'picBow.Top = (Arrows(cmbBow.ListIndex + 1).Pic * 32) * -1
End Sub

Private Sub cmdOk_Click()
    Call ItemEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ItemEditorCancel
End Sub

Private Sub cmbType_Click()
    Call NetFra
    
    CtArme.Visible = False
    FramePD.Visible = False
    CheckEmpi.Enabled = True
    fraBow.Visible = False
    SSTab1.Width = 535
    frmItemEditor.Width = 8145
    
    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            Label3.Caption = "Dommage :"
            Label3.ToolTipText = "Dommage infliger par l'objet"
            lblStrength.ToolTipText = "Dommage infliger par l'objet"
            fraBow.Visible = True
            CtArme.Visible = True
        Else
            Label3.Caption = "Défense :"
            Label3.ToolTipText = "Défense de l'objet"
            lblStrength.ToolTipText = "Défense de l'objet"
        End If
        fraEquipment.Visible = True
        fraAttributes.Visible = True
        FramePD.Visible = True
        
        CheckEmpi.value = 0
        CheckEmpi.Enabled = False
        SSTab1.Width = 755
        frmItemEditor.Width = 11445
    End If
        
    If (cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Call NetFra
        fraVitals.Visible = True
    ElseIf (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Call NetFra
        fraSpell.Visible = True
        lblSpellName.Caption = Trim$(Spell(scrlSpell.value).name)
    ElseIf (cmbType.ListIndex = ITEM_TYPE_MONTURE) Then
    On Error Resume Next
        Call NetFra
        framonture.Visible = True
        Call AffSurfPic(DD_SpriteSurf(skin.value), imgmont, 0, 0)
        CheckEmpi.value = 0
        CheckEmpi.Enabled = False
    ElseIf (cmbType.ListIndex = ITEM_TYPE_SCRIPT) Then
        Call NetFra
        fraobjsc.Visible = True
    ElseIf (cmbType.ListIndex = ITEM_TYPE_PET) Then
        Call NetFra
        scrlPet.Max = MAX_PETS
        fraPet.Visible = True
        CheckEmpi.Enabled = False
    End If
End Sub

Private Sub coulr_Click()
cmd.flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then coul.BackColor = cmd.Color: coul.Tag = cmd.Color: txtName.ForeColor = cmd.Color
End Sub

Private Sub Form_Load()
    scrlPD.Max = MAX_DX_PAPERDOLL
    'picItems.Height = 320 * PIC_Y
    Call AffSurfPic(DD_ItemSurf, picSelect, EditorItemX * PIC_X, EditorItemY * PIC_Y)
    Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
    'Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
    'picBow.Picture = LoadPNG(App.Path & "\GFX\arrows.png")
'    picSprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png")
    Picture4.Height = ((PIC_NPC1 * 32) * Screen.TwipsPerPixelY) + 60
    'imgmont.Height = ((PIC_NPC1 * 32) * Screen.TwipsPerPixelY)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ItemEditorCancel
End Sub



Private Sub HScroll1_Change()
    Label30.Caption = HScroll1.value
End Sub

Private Sub HScroll2_Change()
'Call AffSurfPic(DD_PaperDollSurf, picSelect100, 0, HScroll2.value * PIC_Y)
End Sub


Private Sub picPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorItemX = (x \ PIC_X): EditorItemY = (y \ PIC_Y) + VScroll1.value
    Call AffSurfPic(DD_ItemSurf, picSelect, EditorItemX * PIC_X, EditorItemY * PIC_Y)
    'Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub picPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then EditorItemX = (x \ PIC_X): EditorItemY = (y \ PIC_Y) + VScroll1.value
    Call AffSurfPic(DD_ItemSurf, picSelect, EditorItemX * PIC_X, EditorItemY * PIC_Y)
    'Call BitBlt(picSelect.hDC, 0, 0, PIC_X, PIC_Y, picItems.hDC, EditorItemX * PIC_X, EditorItemY * PIC_Y, SRCCOPY)
End Sub

Private Sub scrlAccessReq_Change()
    With scrlAccessReq
        Select Case .value
            Case 0
                Label17.Caption = "Tous"
            Case 1
                Label17.Caption = "Modérateur"
            Case 2
                Label17.Caption = "Mappeur"
            Case 3
                Label17.Caption = "Dévellopeur"
            Case 4
                Label17.Caption = "Administrateur"
            End Select
    End With
End Sub

Private Sub scrlAddDef_Change()
    lblAddDef.Caption = scrlAddDef.value
End Sub

Private Sub scrlAddEXP_Change()
    lblAddEXP.Caption = scrlAddEXP.value & "%"
End Sub

Private Sub scrlAddHP_Change()
    lblAddHP.Caption = scrlAddHP.value
End Sub

Private Sub scrlAddMagi_Change()
    lblAddMagi.Caption = scrlAddMagi.value
End Sub

Private Sub scrlAddMP_Change()
    lblAddMP.Caption = scrlAddMP.value
End Sub

Private Sub scrlAddSP_Change()
    lblAddSP.Caption = scrlAddSP.value
End Sub

Private Sub scrlAddSpeed_Change()
    lblAddSpeed.Caption = scrlAddSpeed.value
End Sub

Private Sub scrlAddStr_Change()
    lblAddStr.Caption = scrlAddStr.value
End Sub

Private Sub scrlAttackSpeed_Change()
    lblAttackSpeed.Caption = scrlAttackSpeed.value & " Millisecondes"
End Sub

Private Sub scrlClassReq_Change()
If scrlClassReq.value = -1 Then
    Label16.Caption = "Aucune"
Else
    If HORS_LIGNE = 1 Then Label16.Caption = scrlClassReq.value & " - " & "Classe" & scrlClassReq.value Else Label16.Caption = scrlClassReq.value & " - " & Trim$(Class(scrlClassReq.value).name)
End If
End Sub

Private Sub scrlDefReq_Change()
    Label12.Caption = scrlDefReq.value
End Sub

Private Sub scrlPD_Change()
    PicPD.Picture = LoadPNG(App.Path & "\GFX\Paperdolls\Paperdolls" & scrlPD.value & ".png")
    PicPD.Height = 64
    Pic.Height = 1020
    lblpaper.Caption = scrlPD.value
End Sub

Private Sub scrlPet_Change()
    lblPetNum.Caption = scrlPet.value
    lblPetNom.Caption = Pets(scrlPet.value).nom
End Sub

Private Sub scrlSexReq_Change()
    If scrlSexReq.value = 2 Then Label35.Caption = "Les 2"
    If scrlSexReq.value = 1 Then Label35.Caption = "Femme"
    If scrlSexReq.value = 0 Then Label35.Caption = "Homme"
End Sub

Private Sub scrlSpeedReq_Change()
    Label13.Caption = scrlSpeedReq.value
End Sub

Private Sub scrlStrReq_Change()
    Label11.Caption = scrlStrReq.value
End Sub

Private Sub scrlVitalMod_Change()
    lblVitalMod.Caption = CStr(scrlVitalMod.value)
End Sub

Private Sub scrlDurability_Change()
    lblDurability.Caption = CStr(scrlDurability.value)
    If CStr(scrlDurability.value) <= 0 Then lblDurability.Caption = "Ind."
End Sub

Private Sub scrlStrength_Change()
    lblStrength.Caption = CStr(scrlStrength.value)
End Sub

Private Sub scrlSpell_Change()
    lblSpellName.Caption = Trim$(Spell(scrlSpell.value).name)
    lblSpell.Caption = CStr(scrlSpell.value)
End Sub

Private Sub skin_Change()
On Error Resume Next
    numskin.Caption = skin.value
    Call AffSurfPic(DD_SpriteSurf(skin.value), imgmont, 0, 0)
End Sub

Private Sub vit_Change()
If vit.value = 5 Or vit.value = 6 Then vit.value = 8
If vit.value = 7 Then vit.value = 4
If vit.value = 3 And VitVal = 4 Then vit.value = 2
If vit.value = 3 And VitVal = 2 Then vit.value = 4
VitVal = vit.value
    nbvit.Caption = vit.value
End Sub

Private Sub VScroll1_Change()
    Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
    'picItems.Top = (VScroll1.value * PIC_Y) * -1
End Sub

Private Sub NetFra()
    fraobjsc.Visible = False
    framonture.Visible = False
    fraVitals.Visible = False
    fraAttributes.Visible = False
    fraEquipment.Visible = False
    fraBow.Visible = False
    fraSpell.Visible = False
    fraPet.Visible = False
End Sub

Private Sub VScroll1_Scroll()
Call AffSurfPic(DD_ItemSurf, picPic, 0, VScroll1.value * PIC_X)
End Sub

Private Sub VScroll2_Change()

End Sub
