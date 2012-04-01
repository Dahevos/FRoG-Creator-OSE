VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmMapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propriétés de la Carte"
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   90
   ClientWidth     =   10080
   ClipControls    =   0   'False
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
   ScaleWidth      =   672
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8325
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Propriétés de la carte"
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   14684
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Général"
      TabPicture(0)   =   "frmMapProperties.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdOk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCancel"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "PNJ"
      TabPicture(1)   =   "frmMapProperties.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmbNpc(14)"
      Tab(1).Control(1)=   "cmbNpc(13)"
      Tab(1).Control(2)=   "cmbNpc(12)"
      Tab(1).Control(3)=   "cmbNpc(11)"
      Tab(1).Control(4)=   "cmbNpc(10)"
      Tab(1).Control(5)=   "cmbNpc(9)"
      Tab(1).Control(6)=   "cmbNpc(8)"
      Tab(1).Control(7)=   "cmbNpc(7)"
      Tab(1).Control(8)=   "cmbNpc(6)"
      Tab(1).Control(9)=   "cmbNpc(5)"
      Tab(1).Control(10)=   "cmbNpc(4)"
      Tab(1).Control(11)=   "cmbNpc(3)"
      Tab(1).Control(12)=   "cmbNpc(2)"
      Tab(1).Control(13)=   "cmbNpc(1)"
      Tab(1).Control(14)=   "cmbNpc(0)"
      Tab(1).Control(15)=   "Command1"
      Tab(1).Control(16)=   "Copy(0)"
      Tab(1).Control(17)=   "Copy(1)"
      Tab(1).Control(18)=   "Copy(2)"
      Tab(1).Control(19)=   "Copy(3)"
      Tab(1).Control(20)=   "Copy(4)"
      Tab(1).Control(21)=   "Copy(5)"
      Tab(1).Control(22)=   "Copy(6)"
      Tab(1).Control(23)=   "Copy(7)"
      Tab(1).Control(24)=   "Copy(8)"
      Tab(1).Control(25)=   "Copy(10)"
      Tab(1).Control(26)=   "Copy(11)"
      Tab(1).Control(27)=   "Copy(12)"
      Tab(1).Control(28)=   "Copy(13)"
      Tab(1).Control(29)=   "Copy(9)"
      Tab(1).Control(30)=   "plus(10)"
      Tab(1).Control(31)=   "plus(14)"
      Tab(1).Control(32)=   "plus(13)"
      Tab(1).Control(33)=   "plus(12)"
      Tab(1).Control(34)=   "plus(11)"
      Tab(1).Control(35)=   "plus(9)"
      Tab(1).Control(36)=   "plus(8)"
      Tab(1).Control(37)=   "plus(7)"
      Tab(1).Control(38)=   "plus(6)"
      Tab(1).Control(39)=   "plus(5)"
      Tab(1).Control(40)=   "plus(4)"
      Tab(1).Control(41)=   "plus(3)"
      Tab(1).Control(42)=   "plus(2)"
      Tab(1).Control(43)=   "plus(1)"
      Tab(1).Control(44)=   "plus(15)"
      Tab(1).Control(45)=   "Command4"
      Tab(1).Control(46)=   "Command5"
      Tab(1).ControlCount=   47
      Begin VB.CommandButton Command5 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   -67080
         TabIndex        =   90
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Valider"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   89
         Top             =   7680
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Panoramas de la carte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   360
         TabIndex        =   75
         ToolTipText     =   "Panorama situé au dessus de la couche frange"
         Top             =   5040
         Width           =   3615
         Begin VB.TextBox PanoInf 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   81
            ToolTipText     =   "Panorama situé entre la couche mask et la couche frange"
            Top             =   480
            Width           =   2385
         End
         Begin VB.TextBox PanoSup 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   40
            TabIndex        =   80
            ToolTipText     =   "Panorama situé au dessus de la couche frange"
            Top             =   1320
            Width           =   2385
         End
         Begin VB.CommandButton ch1 
            Caption         =   "Choisir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   79
            ToolTipText     =   "Jouer la musique sélectionner"
            Top             =   480
            Width           =   960
         End
         Begin VB.CommandButton ch2 
            Caption         =   "Choisir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   78
            ToolTipText     =   "Jouer la musique sélectionner"
            Top             =   1320
            Width           =   960
         End
         Begin VB.CheckBox TSup 
            Caption         =   "Transparence "
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
            Left            =   120
            TabIndex        =   77
            ToolTipText     =   "Si vous avez mit un fond d'une couleur unie pour faire de la transparence comme sur les tiles cochez cette case"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.CheckBox TInf 
            Caption         =   "Transparence "
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
            Left            =   120
            TabIndex        =   76
            ToolTipText     =   "Si vous avez mit un fond d'une couleur unie pour faire de la transparence comme sur les tiles cochez cette case"
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "Panorama inférieur :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1680
         End
         Begin VB.Label Label4 
            Caption         =   "Panorama supérieur :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   82
            Top             =   1080
            Width           =   1680
         End
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   15
         Left            =   -69480
         TabIndex        =   59
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   1
         Left            =   -69480
         TabIndex        =   16
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   2
         Left            =   -69480
         TabIndex        =   20
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   3
         Left            =   -69480
         TabIndex        =   23
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   4
         Left            =   -69480
         TabIndex        =   26
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   5
         Left            =   -69480
         TabIndex        =   29
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   6
         Left            =   -69480
         TabIndex        =   32
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   7
         Left            =   -69480
         TabIndex        =   35
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   8
         Left            =   -69480
         TabIndex        =   38
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   9
         Left            =   -69480
         TabIndex        =   41
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   11
         Left            =   -69480
         TabIndex        =   47
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   12
         Left            =   -69480
         TabIndex        =   50
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   5880
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   13
         Left            =   -69480
         TabIndex        =   53
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   14
         Left            =   -69480
         TabIndex        =   56
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   6840
         Width           =   615
      End
      Begin VB.CommandButton plus 
         Caption         =   "Plus..."
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
         Index           =   10
         Left            =   -69480
         TabIndex        =   44
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   4920
         Width           =   615
      End
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
         Height          =   420
         Left            =   7920
         TabIndex        =   17
         ToolTipText     =   "Annuler les changements"
         Top             =   7560
         Width           =   1440
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
         Height          =   420
         Left            =   6000
         TabIndex        =   14
         ToolTipText     =   "Confirmer les changements"
         Top             =   7560
         Width           =   1440
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   9
         Left            =   -70320
         TabIndex        =   46
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   13
         Left            =   -70320
         TabIndex        =   58
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   12
         Left            =   -70320
         TabIndex        =   55
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   6840
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   11
         Left            =   -70320
         TabIndex        =   52
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   6360
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   10
         Left            =   -70320
         TabIndex        =   49
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   5880
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   8
         Left            =   -70320
         TabIndex        =   43
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   7
         Left            =   -70320
         TabIndex        =   40
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   6
         Left            =   -70320
         TabIndex        =   37
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3960
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   5
         Left            =   -70320
         TabIndex        =   34
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3480
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   4
         Left            =   -70320
         TabIndex        =   31
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   3
         Left            =   -70320
         TabIndex        =   28
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   2
         Left            =   -70320
         TabIndex        =   25
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   1
         Left            =   -70320
         TabIndex        =   22
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton Copy 
         Caption         =   "Copier"
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
         Index           =   0
         Left            =   -70320
         TabIndex        =   19
         ToolTipText     =   "Copier le pnj au dessus de celui là"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Retirer les PNJ de la carte"
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
         Left            =   -73560
         TabIndex        =   60
         ToolTipText     =   "Retirer tout les pnj de la carte"
         Top             =   7800
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Caption         =   "Musique de la carte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   4320
         TabIndex        =   73
         ToolTipText     =   "Musique entendue par les joueurs qui sont sur la carte"
         Top             =   2880
         Width           =   5055
         Begin VB.CommandButton Command3 
            Caption         =   "Stop"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            TabIndex        =   13
            ToolTipText     =   "Arreter la musique sélectionner"
            Top             =   840
            Width           =   1440
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Jouer"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3480
            TabIndex        =   12
            ToolTipText     =   "Jouer la musique sélectionner"
            Top             =   360
            Width           =   1440
         End
         Begin VB.ListBox lstMusic 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3765
            ItemData        =   "frmMapProperties.frx":0038
            Left            =   120
            List            =   "frmMapProperties.frx":003A
            TabIndex        =   11
            Top             =   285
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Téléportation à la déconnexion :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   360
         TabIndex        =   69
         ToolTipText     =   "Vous pouvez l'utiliser pour faire vos donjons par exemple"
         Top             =   3000
         Width           =   3615
         Begin VB.CommandButton collco 
            Caption         =   "Coller les coordonées"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   74
            ToolTipText     =   "Coller les coordonées enregistrées précédement"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtBootMap 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtBootX 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   9
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtBootY 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "0"
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Numéros de la carte :"
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
            Left            =   480
            TabIndex        =   72
            Top             =   360
            Width           =   1650
         End
         Begin VB.Label Label8 
            Caption         =   "Valeur en X :"
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
            Left            =   480
            TabIndex        =   71
            Top             =   690
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Valeur en Y :"
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
            Left            =   480
            TabIndex        =   70
            Top             =   1080
            Width           =   960
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Globale"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   4320
         TabIndex        =   67
         Top             =   840
         Width           =   5085
         Begin VB.HScrollBar prAlpha 
            Enabled         =   0   'False
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   100
            TabIndex        =   88
            Top             =   1560
            Width           =   3135
         End
         Begin VB.ComboBox cmbFog 
            Enabled         =   0   'False
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
            ItemData        =   "frmMapProperties.frx":003C
            Left            =   3600
            List            =   "frmMapProperties.frx":003E
            Style           =   2  'Dropdown List
            TabIndex        =   87
            ToolTipText     =   "Sélectionner le numéros du fichier de brouillard corespondant"
            Top             =   1530
            Width           =   975
         End
         Begin VB.CheckBox chkFog 
            Caption         =   "Brouillard"
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
            Left            =   240
            TabIndex        =   84
            ToolTipText     =   "Active/Désactive le brouillard sur la carte sélectionnée"
            Top             =   1000
            Width           =   975
         End
         Begin VB.ComboBox cmbMoral 
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
            ItemData        =   "frmMapProperties.frx":0040
            Left            =   240
            List            =   "frmMapProperties.frx":004D
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Sélectionner une spécialité (Pvp : joueurs contre joueurs)"
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label5 
            Caption         =   "Numéros :"
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
            Left            =   3600
            TabIndex        =   86
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label prFog 
            Caption         =   "Pourcentage de transparence :"
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
            Left            =   240
            TabIndex        =   85
            Top             =   1320
            Width           =   2895
         End
         Begin VB.Label Label1 
            Caption         =   "Spécialité de la carte :"
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
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Téléportation sur les bords de la Carte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   360
         TabIndex        =   62
         Top             =   840
         Width           =   3615
         Begin VB.CheckBox Ctraversable 
            Caption         =   "Joueur Traversable"
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
            Left            =   1320
            TabIndex        =   93
            ToolTipText     =   "La nuit ne tombera pas sur cette carte si la case est cochée"
            Top             =   1800
            Width           =   2055
         End
         Begin VB.CheckBox cPetView 
            Caption         =   "Cacher Famillier sur la map"
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
            Left            =   1320
            TabIndex        =   92
            ToolTipText     =   "La nuit ne tombera pas sur cette carte si la case est cochée"
            Top             =   1560
            Width           =   2200
         End
         Begin VB.CheckBox CGuild 
            Caption         =   "Guilde Seule"
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
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "La nuit ne tombera pas sur cette carte si la case est cochée"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox chkIndoors 
            Caption         =   "Intérieur"
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
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "La nuit ne tombera pas sur cette carte si la case est cochée"
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox txtLeft 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2070
            TabIndex        =   5
            Text            =   "0"
            ToolTipText     =   "Numéros de la carte où les joueurs seront téléporter"
            Top             =   1260
            Width           =   855
         End
         Begin VB.TextBox txtDown 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2070
            TabIndex        =   4
            Text            =   "0"
            ToolTipText     =   "Numéros de la carte où les joueurs seront téléporter"
            Top             =   975
            Width           =   855
         End
         Begin VB.TextBox txtRight 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2070
            TabIndex        =   3
            Text            =   "0"
            ToolTipText     =   "Numéros de la carte où les joueurs seront téléporter"
            Top             =   690
            Width           =   855
         End
         Begin VB.TextBox txtUp 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2070
            TabIndex        =   2
            Text            =   "0"
            ToolTipText     =   "Numéros de la carte où les joueurs seront téléporter"
            Top             =   405
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Ouest(Gauche) :"
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
            Left            =   600
            TabIndex        =   66
            Top             =   1275
            Width           =   1365
         End
         Begin VB.Label Label15 
            Caption         =   "Sud(Bas) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   65
            Top             =   1005
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Est(Droite) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   600
            TabIndex        =   64
            Top             =   705
            Width           =   1095
         End
         Begin VB.Label Label14 
            Caption         =   "Nord(Haut) :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   600
            TabIndex        =   63
            Top             =   405
            Width           =   1020
         End
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         MaxLength       =   40
         TabIndex        =   1
         ToolTipText     =   "Ecrivez le nom désirer pour la carte ici"
         Top             =   360
         Width           =   7665
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   0
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   15
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   600
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   1
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   2
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   3
         ItemData        =   "frmMapProperties.frx":007E
         Left            =   -74520
         List            =   "frmMapProperties.frx":0080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   2040
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   4
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   2520
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   5
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   30
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   3000
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   6
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   33
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   3480
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   7
         ItemData        =   "frmMapProperties.frx":0082
         Left            =   -74520
         List            =   "frmMapProperties.frx":0084
         Style           =   2  'Dropdown List
         TabIndex        =   36
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   3960
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   8
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   39
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   4440
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   9
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   42
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   4920
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   10
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   45
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   5400
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   11
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   48
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   5880
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   12
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   51
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   6360
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   13
         ItemData        =   "frmMapProperties.frx":0086
         Left            =   -74520
         List            =   "frmMapProperties.frx":0088
         Style           =   2  'Dropdown List
         TabIndex        =   54
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   6840
         Width           =   4095
      End
      Begin VB.ComboBox cmbNpc 
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
         Index           =   14
         Left            =   -74520
         Style           =   2  'Dropdown List
         TabIndex        =   57
         ToolTipText     =   "Sélectionner un pnj"
         Top             =   7320
         Width           =   4095
      End
      Begin VB.Label Label13 
         Caption         =   "Nom de la carte :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   61
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ch1_Click()
frmPanorama.lstPano.Tag = 0
ListPanorama (App.Path & "\GFX\")
frmPanorama.Show vbModal, Me
End Sub

Private Sub ch2_Click()
frmPanorama.lstPano.Tag = 1
ListPanorama (App.Path & "\GFX\")
frmPanorama.Show vbModal, Me
End Sub

Private Sub chkFog_Click()
    prAlpha.Enabled = chkFog.value
    cmbFog.Enabled = chkFog.value
    If chkFog.value = 0 Then
        prAlpha.value = 0
        cmbFog.ListIndex = 0
    End If
End Sub

Private Sub collco_Click()
txtBootMap.Text = CoordM
txtBootX.Text = CoordX
txtBootY.Text = CoordY
End Sub

Private Sub Command1_Click()
Dim i As Long
For i = 1 To MAX_MAP_NPCS
    cmbNpc(i - 1).ListIndex = 0
Next i
End Sub

Private Sub Command2_Click()
Call StopMidi
Call PlayMidi(lstMusic.Text)
End Sub

Private Sub Command3_Click()
Call StopMidi
End Sub

Private Sub Command4_Click()
Dim x As Long, y As Long, i As Long

    Map(Player(MyIndex).Map).name = txtName.Text
    Map(Player(MyIndex).Map).Up = Val(txtUp.Text)
    Map(Player(MyIndex).Map).Down = Val(txtDown.Text)
    Map(Player(MyIndex).Map).Left = Val(txtLeft.Text)
    Map(Player(MyIndex).Map).Right = Val(txtRight.Text)
    Map(Player(MyIndex).Map).Moral = cmbMoral.ListIndex
    Map(Player(MyIndex).Map).Music = lstMusic.Text
    Map(Player(MyIndex).Map).BootMap = Val(txtBootMap.Text)
    Map(Player(MyIndex).Map).BootX = Val(txtBootX.Text)
    Map(Player(MyIndex).Map).BootY = Val(txtBootY.Text)
    Map(Player(MyIndex).Map).Indoors = Val(chkIndoors.value)
    Map(Player(MyIndex).Map).PanoInf = PanoInf.Text
    Map(Player(MyIndex).Map).TranInf = Val(TInf.value)
    Map(Player(MyIndex).Map).PanoSup = PanoSup.Text
    Map(Player(MyIndex).Map).TranSup = Val(TSup.value)
    Map(Player(MyIndex).Map).Fog = Val(cmbFog.Text)
    Map(Player(MyIndex).Map).FogAlpha = 100 - Val(prAlpha.value)
    Map(Player(MyIndex).Map).guildSoloView = Val(CGuild.value)
    Map(Player(MyIndex).Map).petView = Val(cPetView.value)
    Map(Player(MyIndex).Map).traversable = Val(Ctraversable.value)
    
    For i = 1 To MAX_MAP_NPCS
        Map(Player(MyIndex).Map).Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Call InitPano(Player(MyIndex).Map)
    Call InitNightAndFog(Player(MyIndex).Map)
    Call StopMidi
    InProprieter = False
    Me.Hide
    frmMirage.Show
End Sub

Private Sub Command5_Click()
InProprieter = False
Call StopMidi
Me.Hide
End Sub

Private Sub Copy_Click(Index As Integer)
    cmbNpc(Index + 1).ListIndex = cmbNpc(Index).ListIndex
End Sub

Public Sub InitMPr()
frmMapProperties.Caption = frmMapProperties.Caption & Player(MyIndex).Map & " Mettez votre souris sur un élément pour plus de détails."
Dim x As Long, y As Long, i As Long

    txtName.Text = Trim$(Map(Player(MyIndex).Map).name)
    txtUp.Text = CStr(Map(Player(MyIndex).Map).Up)
    txtDown.Text = CStr(Map(Player(MyIndex).Map).Down)
    txtLeft.Text = CStr(Map(Player(MyIndex).Map).Left)
    txtRight.Text = CStr(Map(Player(MyIndex).Map).Right)
    cmbMoral.ListIndex = Map(Player(MyIndex).Map).Moral
    txtBootMap.Text = CStr(Map(Player(MyIndex).Map).BootMap)
    txtBootX.Text = CStr(Map(Player(MyIndex).Map).BootX)
    txtBootY.Text = CStr(Map(Player(MyIndex).Map).BootY)
    ListMusic (App.Path & "\Music\")
    lstMusic = Trim$(Map(Player(MyIndex).Map).Music)
    lstMusic.Text = Trim$(Map(Player(MyIndex).Map).Music)
    chkIndoors.value = CStr(Map(Player(MyIndex).Map).Indoors)
    CGuild.value = CStr(Map(Player(MyIndex).Map).guildSoloView)
    cPetView.value = CStr(Map(Player(MyIndex).Map).petView)
    Ctraversable.value = CStr(Map(Player(MyIndex).Map).traversable)
    PanoInf.Text = Trim$(Map(Player(MyIndex).Map).PanoInf)
    TInf.value = Val(Map(Player(MyIndex).Map).TranInf)
    PanoSup.Text = Trim$(Map(Player(MyIndex).Map).PanoSup)
    TSup.value = Val(Map(Player(MyIndex).Map).TranSup)
    ListFog (App.Path & "\GFX\")
    cmbFog.ListIndex = LstID(Map(Player(MyIndex).Map).Fog)
    prAlpha.value = 100 - Map(Player(MyIndex).Map).FogAlpha
    If Map(Player(MyIndex).Map).Fog <> 0 Then chkFog.value = 1
    prAlpha.Enabled = chkFog.value
    cmbFog.Enabled = chkFog.value
    
    For x = 1 To MAX_MAP_NPCS
        cmbNpc(x - 1).Clear
        cmbNpc(x - 1).AddItem "Pas de PNJ"
    Next x
    
    For y = 1 To MAX_NPCS
        For x = 1 To MAX_MAP_NPCS
            cmbNpc(x - 1).AddItem y & " : " & Trim$(Npc(y).name)
        Next x
    Next y
    
    For i = 1 To MAX_MAP_NPCS
        cmbNpc(i - 1).ListIndex = Map(Player(MyIndex).Map).Npc(i)
    Next i
    
    Call StopMidi
End Sub

Private Sub Form_Load()
Dim x As Long, y As Long, i As Long
    Call StopMidi
End Sub

Private Sub cmdOk_Click()
Dim x As Long, y As Long, i As Long

    Map(Player(MyIndex).Map).name = txtName.Text
    Map(Player(MyIndex).Map).Up = Val(txtUp.Text)
    Map(Player(MyIndex).Map).Down = Val(txtDown.Text)
    Map(Player(MyIndex).Map).Left = Val(txtLeft.Text)
    Map(Player(MyIndex).Map).Right = Val(txtRight.Text)
    Map(Player(MyIndex).Map).Moral = cmbMoral.ListIndex
    Map(Player(MyIndex).Map).Music = lstMusic.Text
    Map(Player(MyIndex).Map).BootMap = Val(txtBootMap.Text)
    Map(Player(MyIndex).Map).BootX = Val(txtBootX.Text)
    Map(Player(MyIndex).Map).BootY = Val(txtBootY.Text)
    Map(Player(MyIndex).Map).Indoors = Val(chkIndoors.value)
    Map(Player(MyIndex).Map).PanoInf = PanoInf.Text
    Map(Player(MyIndex).Map).TranInf = Val(TInf.value)
    Map(Player(MyIndex).Map).PanoSup = PanoSup.Text
    Map(Player(MyIndex).Map).TranSup = Val(TSup.value)
    Map(Player(MyIndex).Map).Fog = Val(cmbFog.Text)
    Map(Player(MyIndex).Map).FogAlpha = 100 - Val(prAlpha.value)
    Map(Player(MyIndex).Map).guildSoloView = Val(CGuild.value)
    Map(Player(MyIndex).Map).petView = Val(cPetView.value)
    Map(Player(MyIndex).Map).traversable = Val(Ctraversable.value)
    
    For i = 1 To MAX_MAP_NPCS
        Map(Player(MyIndex).Map).Npc(i) = cmbNpc(i - 1).ListIndex
    Next i
    
    Call InitPano(Player(MyIndex).Map)
    Call InitNightAndFog(Player(MyIndex).Map)
    Call StopMidi
    InProprieter = False
    Me.Hide
    frmMirage.Show
End Sub

Private Sub cmdCancel_Click()
InProprieter = False
Call StopMidi
Me.Hide
End Sub

Private Sub Form_Terminate()
Me.Hide
If Trim$(Map(Player(MyIndex).Map).Music) <> vbNullString Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
frmMirage.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Trim$(Map(Player(MyIndex).Map).Music) <> vbNullString Then Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
frmMirage.Show
End Sub

Private Sub PanoInf_Click()
frmPanorama.lstPano.Tag = 0
ListPanorama (App.Path & "\GFX\")
frmPanorama.Show vbModal, Me
End Sub

Private Sub PanoInf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then PanoInf.Text = vbNullString
End Sub

Private Sub PanoInf_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then PanoInf.Text = vbNullString
End Sub

Private Sub PanoSup_Click()
frmPanorama.lstPano.Tag = 1
ListPanorama (App.Path & "\GFX\")
frmPanorama.Show vbModal, Me
End Sub

Private Sub PanoSup_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyDelete Then PanoInf.Text = vbNullString
End Sub

Private Sub plus_Click(Index As Integer)
InMouvEditor = True
EditorMouvIndex = Index
frmCpnjmouv.Caption = "Editer les mouvement du PNJ" & Index & " de la Carte" & Player(MyIndex).Map
frmCpnjmouv.Show
frmCpnjmouv.SetFocus
End Sub

Private Sub prAlpha_Change()
    prFog.Caption = "Pourcentage de transparence : " & prAlpha.value & "%"
End Sub

Private Sub prAlpha_Scroll()
    prFog.Caption = "Pourcentage de transparence : " & prAlpha.value & "%"
End Sub

Private Sub txtBootMap_GotFocus()
txtBootMap.SelStart = 0
txtBootMap.SelLength = Len(txtBootMap)
End Sub

Private Sub txtBootX_GotFocus()
txtBootX.SelStart = 0
txtBootX.SelLength = Len(txtBootX)
End Sub

Private Sub txtBootY_GotFocus()
txtBootY.SelStart = 0
txtBootY.SelLength = Len(txtBootY)
End Sub

Private Sub txtDown_GotFocus()
txtDown.SelStart = 0
txtDown.SelLength = Len(txtDown)
End Sub

Private Sub txtLeft_GotFocus()
txtLeft.SelStart = 0
txtLeft.SelLength = Len(txtLeft)
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName)
End Sub

Private Sub txtRight_GotFocus()
txtRight.SelStart = 0
txtRight.SelLength = Len(txtRight)
End Sub

Private Sub txtUp_GotFocus()
txtRight.SelStart = 0
txtRight.SelLength = Len(txtRight)
End Sub

Private Function LstID(ByVal tx As Long) As Long
On Error Resume Next
Dim i As Long
LstID = 0
    For i = 0 To cmbFog.ListCount
        If Val(cmbFog.List(i)) = tx Then LstID = i: Exit For
    Next i
End Function
