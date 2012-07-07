VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmadmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panneau d'administration"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "frmadmin"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnclose 
      Caption         =   "Fermer le panneau"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   4920
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   9128
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   353
      TabMaxWidth     =   1940
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Commande"
      TabPicture(0)   =   "frmadmin.frx":17D2A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Commande"
      TabPicture(1)   =   "frmadmin.frx":17D46
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Aide"
      TabPicture(2)   =   "frmadmin.frx":17D62
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame11 
         Caption         =   "Environnement"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   2400
         TabIndex        =   60
         Top             =   2760
         Width           =   2055
         Begin VB.TextBox motd 
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
            TabIndex        =   63
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Changer mot de bienvenue"
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
            TabIndex        =   62
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Jour / Nuit"
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
            TabIndex        =   61
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Mot de bienvenue:"
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
            TabIndex        =   64
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Commande du sprite"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -72600
         TabIndex        =   55
         Top             =   2160
         Width           =   2055
         Begin VB.CommandButton btnPlayerSprite 
            Caption         =   "Changer sprite du joueur"
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
            TabIndex        =   56
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtSprite 
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
            TabIndex        =   58
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton btnSprite 
            Caption         =   "Changer votre sprite"
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
            TabIndex        =   57
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "Numéro du sprite:"
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
            TabIndex        =   59
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Chagement de Stats"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   -74880
         TabIndex        =   42
         Top             =   1440
         Width           =   2055
         Begin VB.CommandButton Command11 
            Caption         =   "Changer le PM"
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
            TabIndex        =   53
            Top             =   2400
            Width           =   1815
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Changer les PV"
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
            TabIndex        =   52
            Top             =   2160
            Width           =   1815
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Changer les points"
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
            TabIndex        =   54
            Top             =   1920
            Width           =   1815
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Changer l'expérience"
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
            TabIndex        =   44
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Changer le niveau"
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
            TabIndex        =   45
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Changer le PK"
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
            TabIndex        =   46
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Changer la magie"
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
            TabIndex        =   47
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Changer la vitesse"
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
            TabIndex        =   48
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Changer la défense"
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
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Changer la force"
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
            TabIndex        =   50
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtValeur 
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "Nouvelle valeur:"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   1815
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Cible de la Commande"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   -74880
         TabIndex        =   39
         Top             =   600
         Width           =   2055
         Begin VB.TextBox txtplayer2 
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label17 
            Caption         =   "Nom du Joueur concerné:"
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
            TabIndex        =   41
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Commande de Nom"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -72600
         TabIndex        =   34
         Top             =   600
         Width           =   2055
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
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1815
         End
         Begin VB.CommandButton btnname 
            Caption         =   "Changer le nom du joueur"
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
            TabIndex        =   36
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton btnyname 
            Caption         =   "Changer votre nom"
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
            TabIndex        =   35
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Nouveau nom:"
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
            TabIndex        =   38
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Conseil"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   32
         Top             =   3240
         Width           =   4335
         Begin VB.TextBox Text1 
            Height          =   1095
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Text            =   "frmadmin.frx":17D7E
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Légende :"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   22
         Top             =   600
         Width           =   4335
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   2295
            Left            =   120
            Picture         =   "frmadmin.frx":17E39
            ScaleHeight     =   2235
            ScaleWidth      =   4035
            TabIndex        =   23
            Top             =   240
            Width           =   4095
            Begin VB.Label Label8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "5 : Administrateur en chef (Compte Suprême)"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   2040
               Width           =   3495
            End
            Begin VB.Label Label9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "4 : Administrateur (Compte supérieur)"
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   1800
               Width           =   3015
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFFFFF&
               Caption         =   "3 : Développeur (Compte supérieur)"
               Height          =   255
               Left            =   120
               TabIndex        =   29
               Top             =   1560
               Width           =   3015
            End
            Begin VB.Label Label11 
               BackColor       =   &H00FFFFFF&
               Caption         =   "2 : Mappeur ( Compte supérieur)"
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   1320
               Width           =   3015
            End
            Begin VB.Label Label12 
               BackColor       =   &H00FFFFFF&
               Caption         =   "1 : Modérateur (Compte supérieur)"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   1080
               Width           =   3015
            End
            Begin VB.Label Label13 
               BackColor       =   &H00FFFFFF&
               Caption         =   "0 : Joueur (Compte normal)"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   840
               Width           =   3255
            End
            Begin VB.Label Label14 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Voiçi les différents niveau d'access possible."
               Height          =   255
               Left            =   120
               TabIndex        =   25
               Top             =   600
               Width           =   3855
            End
            Begin VB.Label Label15 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Les accès.."
               Height          =   255
               Left            =   540
               TabIndex        =   24
               Top             =   180
               Width           =   2655
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Cible de la Commande"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   2055
         Begin VB.TextBox txtPlayer 
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label1 
            Caption         =   "Nom du Joueur concerné:"
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
            TabIndex        =   18
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Commande de Maps"
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
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   2055
         Begin VB.CommandButton Command15 
            Caption         =   "Menu de téléportation"
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
            TabIndex        =   65
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton btnWarpto 
            Caption         =   "Téléporter à .."
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
            TabIndex        =   11
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton btnRespawn 
            Caption         =   "Réinitialiser"
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
            TabIndex        =   12
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton btnLOC 
            Caption         =   "Localisation"
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
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox txtMap 
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
            TabIndex        =   14
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Numéro de la Map:"
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
            TabIndex        =   15
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Commande Maitre de Jeu"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
         Begin VB.CommandButton Command2 
            Caption         =   "Informations sur le joueur"
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
            TabIndex        =   20
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CommandButton btnKick 
            Caption         =   "Déconnecter"
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
            TabIndex        =   5
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CommandButton btnSetAccess 
            Caption         =   "Changer les access"
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
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtAccess 
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
            TabIndex        =   8
            Top             =   2040
            Width           =   1815
         End
         Begin VB.CommandButton btnBan 
            Caption         =   "Bannir"
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
            TabIndex        =   7
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton btnWarptome 
            Caption         =   "Téléportez-le à moi"
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
            TabIndex        =   66
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton btnWarpMeTo 
            Caption         =   "Téléporter moi à.."
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
            TabIndex        =   6
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Valeur de l'access:"
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
            TabIndex        =   9
            Top             =   1800
            Width           =   1695
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Panneau d'administration 0.4"
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
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Panneau d'administration"
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
         Left            =   -74880
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2280
         X2              =   2280
         Y1              =   360
         Y2              =   4995
      End
      Begin VB.Label Label5 
         Caption         =   "Panneau d'administration"
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
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnname_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtName.Text) <> vbNullString Then Call SendSetPlayerName(Trim$(txtPlayer.Text), Trim$(txtName.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub btnPlayerSprite_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtSprite.Text) <> vbNullString Then Call SendSetPlayerSprite(Trim$(txtPlayer.Text), Trim$(txtSprite.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub btnBan_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call SendBan(Trim$(txtPlayer.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnkick_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MONITER Then
    Call SendKick(Trim$(txtPlayer.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnLOC_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call SendRequestLocation
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub
Private Sub btnRespawn_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call SendMapRespawn
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub
Private Sub btnWarpmeTo_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call WarpMeTo(Trim$(txtPlayer.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnWarpto_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call WarpTo(Val(txtMap.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnWarptome_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call WarpToMe(Trim$(txtPlayer.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnclose_Click()
frmadmin.Visible = False
End Sub

Private Sub btnSprite_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call SendSetSprite(Val(txtSprite.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnSetAccess_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
    Call SendSetAccess(Trim$(txtPlayer.Text), Trim$(txtAccess.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub btnyname_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    If Len(txtName.Text) > 2 Then
    Call SendSetName(Trim$(txtName.Text))
    Else
    MsgBox ("Le nombre de caractères du nom doit être supérieur à 2 caractères")
    End If
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub Command10_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerMaxPv(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command11_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerMaxPm(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command12_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerPoint(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command13_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    If GameTime = TIME_DAY Then
        GameTime = TIME_NIGHT
        Call InitNightAndFog(GetPlayerMap(MyIndex))
    Else
        GameTime = TIME_DAY
    End If
    
    Call SendGameTime
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub Command14_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call SendMOTDChange(Trim$(motd.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub Command15_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
    Call SendData("mapreport" & END_CHAR)
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub Command2_Click()
If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
    Call SendPlayerInfoRequest(Trim$(txtPlayer.Text))
Else
    Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
End If
End Sub

Private Sub Command3_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerDef(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command4_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerVit(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command5_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerstr(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command6_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerMagi(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command7_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerPk(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command8_Click()
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString And Val(txtValeur.Text) < MAX_LEVEL Then Call SendSetPlayerNiveau(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
End Sub

Private Sub Command9_Click()
On Error GoTo er:
    If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
        If Trim$(txtplayer2.Text) <> vbNullString Then
            If Trim$(txtValeur.Text) <> vbNullString Then Call SendSetPlayerExp(Trim$(txtplayer2.Text), Trim$(txtValeur.Text))
        End If
    Else
        Call AddText("Vous n'êtes pas autorisé à faire cette action.", BrightRed)
    End If
Exit Sub
er:
MsgBox "Valeur trop grande."
End Sub

Private Sub motd_Change()
motd.SelStart = 0
motd.SelLength = Len(motd.Text)
End Sub

Private Sub txtAccess_Change()
txtAccess.SelStart = 0
txtAccess.SelLength = Len(txtAccess.Text)
End Sub

Private Sub txtMap_GotFocus()
txtMap.SelStart = 0
txtMap.SelLength = Len(txtMap.Text)
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtPlayer_GotFocus()
txtPlayer.SelStart = 0
txtPlayer.SelLength = Len(txtPlayer.Text)
End Sub

Private Sub txtplayer2_GotFocus()
txtplayer2.SelStart = 0
txtplayer2.SelLength = Len(txtplayer2.Text)
End Sub

Private Sub txtSprite_GotFocus()
txtSprite.SelStart = 0
txtSprite.SelLength = Len(txtSprite.Text)
End Sub

Private Sub txtValeur_GotFocus()
txtValeur.SelStart = 0
txtValeur.SelLength = Len(txtValeur.Text)
End Sub
