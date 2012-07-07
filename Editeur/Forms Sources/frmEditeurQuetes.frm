VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditeurQuetes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Quêtes"
   ClientHeight    =   8730
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   397
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Quêtes"
      TabPicture(0)   =   "frmEditeurQuetes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frtp(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frtp(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frtp(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frtp(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "frtp(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "frtp(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "description"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "tpe"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "nom"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "reponse"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "anul"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ok"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "reconpense"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      Begin VB.Frame reconpense 
         Caption         =   "Récompense"
         Height          =   3375
         Left            =   4560
         TabIndex        =   63
         Top             =   4560
         Width           =   6855
         Begin VB.TextBox rexp 
            Height          =   285
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   75
            ToolTipText     =   "Points d'expérience gagnés par le joueur à la fin de la quête"
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox cases 
            Height          =   285
            Left            =   3000
            TabIndex        =   73
            ToolTipText     =   "Numéros de la case qui va étre éxécutée a la fin de la quête"
            Top             =   3000
            Width           =   3495
         End
         Begin VB.HScrollBar rq3 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            TabIndex        =   31
            Top             =   2640
            Width           =   3495
         End
         Begin VB.HScrollBar ro3 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            Min             =   1
            TabIndex        =   30
            Top             =   2280
            Value           =   1
            Width           =   3495
         End
         Begin VB.HScrollBar rq2 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            TabIndex        =   29
            Top             =   1920
            Width           =   3495
         End
         Begin VB.HScrollBar ro2 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            Min             =   1
            TabIndex        =   28
            Top             =   1560
            Value           =   1
            Width           =   3495
         End
         Begin VB.HScrollBar rq1 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            TabIndex        =   27
            Top             =   1200
            Width           =   3495
         End
         Begin VB.HScrollBar ro1 
            Height          =   255
            LargeChange     =   10
            Left            =   3000
            Max             =   1000
            Min             =   1
            TabIndex        =   26
            Top             =   840
            Value           =   1
            Width           =   3495
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Case scriptée à éxécuter :"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   3000
            Width           =   1845
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Quantité de l'objet3 :"
            Height          =   195
            Left            =   240
            TabIndex        =   70
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Numéro de l'objet3 :"
            Height          =   195
            Left            =   240
            TabIndex        =   69
            Top             =   2280
            Width           =   1410
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Quantité de l'objet2 :"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   1920
            Width           =   1455
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Numéro de l'objet2 :"
            Height          =   195
            Left            =   240
            TabIndex        =   67
            Top             =   1560
            Width           =   1410
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Quantité de l'objet1 :"
            Height          =   195
            Left            =   240
            TabIndex        =   66
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Numéro de l'objet1 :"
            Height          =   195
            Left            =   240
            TabIndex        =   65
            Top             =   840
            Width           =   1410
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Points d'expérience gagner :"
            Height          =   195
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   2010
         End
      End
      Begin VB.CommandButton ok 
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
         Left            =   3960
         TabIndex        =   32
         ToolTipText     =   "Quitte la fenêtre d'édition et enregistre l'objet"
         Top             =   8040
         Width           =   1455
      End
      Begin VB.CommandButton anul 
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
         TabIndex        =   33
         ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer l'objet"
         Top             =   8040
         Width           =   1455
      End
      Begin VB.TextBox reponse 
         Height          =   2595
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "C'est ce que dirat le PNJ qui donnent la quête quand le joueure reviendra le voir a la fin de sa quête"
         Top             =   5280
         Width           =   4095
      End
      Begin VB.TextBox nom 
         Height          =   285
         Left            =   120
         MaxLength       =   40
         TabIndex        =   1
         ToolTipText     =   "Nom de la quête"
         Top             =   840
         Width           =   4095
      End
      Begin VB.ComboBox tpe 
         Height          =   315
         ItemData        =   "frmEditeurQuetes.frx":001C
         Left            =   120
         List            =   "frmEditeurQuetes.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Type de la quête"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox description 
         Height          =   2325
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "C'est ce que les joueurs doivent faire pour terminer la quête"
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   2175
         Index           =   4
         Left            =   4560
         TabIndex        =   44
         Top             =   480
         Width           =   6855
         Begin VB.HScrollBar nbt 
            Height          =   255
            LargeChange     =   10
            Left            =   3360
            Max             =   100
            TabIndex        =   18
            Top             =   1560
            Width           =   3255
         End
         Begin VB.HScrollBar indpnj 
            Height          =   255
            Left            =   240
            Max             =   15
            Min             =   1
            TabIndex        =   15
            Top             =   720
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   4
            LargeChange     =   60
            Left            =   3360
            Max             =   1800
            TabIndex        =   16
            Top             =   720
            Width           =   3255
         End
         Begin VB.HScrollBar numopnj 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            Min             =   1
            TabIndex        =   17
            Top             =   1560
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de fois qu'il faut le tuer :"
            Height          =   195
            Left            =   3360
            TabIndex        =   56
            Top             =   1200
            Width           =   2265
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Index du PNJ (pour la quête) :"
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   2115
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réaliser la quête :"
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   52
            Top             =   360
            Width           =   2085
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Numéro du PNJ :"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   2175
         Index           =   1
         Left            =   4560
         TabIndex        =   34
         Top             =   480
         Width           =   6855
         Begin VB.HScrollBar quant 
            Height          =   255
            LargeChange     =   10
            Left            =   3360
            Max             =   500
            TabIndex        =   8
            Top             =   1560
            Value           =   1
            Width           =   3255
         End
         Begin VB.HScrollBar numo 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            Min             =   1
            TabIndex        =   7
            Top             =   1560
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   1
            LargeChange     =   60
            Left            =   3360
            Max             =   1800
            TabIndex        =   6
            Top             =   720
            Width           =   3255
         End
         Begin VB.HScrollBar indo 
            Height          =   255
            Left            =   240
            Max             =   15
            Min             =   1
            TabIndex        =   5
            Top             =   720
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Quantité à ramassé :"
            Height          =   195
            Left            =   3360
            TabIndex        =   41
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Numéros de l'objet :"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réalisée la Quete :"
            Height          =   195
            Index           =   1
            Left            =   3360
            TabIndex        =   39
            Top             =   360
            Width           =   2160
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Index de l'objet (pour la Quete) :"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Width           =   2250
         End
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   1335
         Index           =   3
         Left            =   4560
         TabIndex        =   43
         Top             =   480
         Width           =   6855
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   3
            LargeChange     =   60
            Left            =   3240
            Max             =   1800
            TabIndex        =   14
            Top             =   720
            Width           =   3255
         End
         Begin VB.HScrollBar numepnj 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            Min             =   1
            TabIndex        =   13
            Top             =   720
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réalisée la Quete :"
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   55
            Top             =   360
            Width           =   2160
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Numéros du PNJ :"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   3975
         Index           =   2
         Left            =   4560
         TabIndex        =   42
         Top             =   480
         Width           =   6855
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   2
            LargeChange     =   60
            Left            =   3240
            Max             =   1800
            TabIndex        =   10
            Top             =   600
            Width           =   3255
         End
         Begin VB.TextBox reppnj 
            Height          =   1605
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2160
            Width           =   6375
         End
         Begin VB.HScrollBar numpnj 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            Min             =   1
            TabIndex        =   11
            Top             =   1440
            Value           =   1
            Width           =   2775
         End
         Begin VB.HScrollBar numod 
            Height          =   255
            LargeChange     =   10
            Left            =   240
            Max             =   1000
            Min             =   1
            TabIndex        =   9
            Top             =   600
            Value           =   1
            Width           =   2775
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réalisée la Quete :"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   54
            Top             =   240
            Width           =   2160
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Réponse du PNJ :"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   1800
            Width           =   1305
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Numéros du PNJ :"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Numéros de l'objet donné :"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   240
            Width           =   1890
         End
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   2295
         Index           =   5
         Left            =   4560
         TabIndex        =   45
         Top             =   480
         Width           =   6855
         Begin VB.CommandButton collco 
            Caption         =   "Coller les coordonées"
            Height          =   255
            Left            =   3720
            TabIndex        =   74
            ToolTipText     =   "Coller les coordonées enregistrées précédement"
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox carted 
            Height          =   285
            Left            =   3120
            MaxLength       =   4
            TabIndex        =   20
            Text            =   "0"
            ToolTipText     =   "Coordonner X"
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton defxy 
            Caption         =   "Définir..."
            Height          =   255
            Left            =   3720
            TabIndex        =   23
            ToolTipText     =   "Définir les Coordonées sur la carte"
            Top             =   1410
            Width           =   735
         End
         Begin VB.TextBox xd 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "0"
            ToolTipText     =   "Coordonner X"
            Top             =   1410
            Width           =   375
         End
         Begin VB.TextBox yd 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   22
            Text            =   "0"
            ToolTipText     =   "Coordonner Y"
            Top             =   1800
            Width           =   375
         End
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   5
            LargeChange     =   60
            Left            =   240
            Max             =   1800
            TabIndex        =   19
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Numeros de la carte du donjon :"
            Height          =   195
            Left            =   240
            TabIndex        =   71
            Top             =   1125
            Width           =   2265
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Coordonnée en Y de la fin du donjon :"
            Height          =   195
            Left            =   240
            TabIndex        =   59
            Top             =   1845
            Width           =   2685
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Coordonnée en X de la fin du donjon :"
            Height          =   195
            Left            =   240
            TabIndex        =   58
            Top             =   1455
            Width           =   2685
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réalisée la Quete :"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   2160
         End
      End
      Begin VB.Frame frtp 
         Caption         =   "Caractéristiques"
         Height          =   2175
         Index           =   6
         Left            =   4560
         TabIndex        =   46
         Top             =   480
         Width           =   6855
         Begin VB.HScrollBar nbxp 
            Height          =   255
            LargeChange     =   100
            Left            =   240
            Max             =   30000
            Min             =   1
            TabIndex        =   25
            Top             =   1560
            Value           =   1
            Width           =   6255
         End
         Begin VB.HScrollBar tempr 
            Height          =   255
            Index           =   6
            LargeChange     =   60
            Left            =   240
            Max             =   1800
            TabIndex        =   24
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de points d'experiences a gagner :"
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   1200
            Width           =   3030
         End
         Begin VB.Label tp 
            AutoSize        =   -1  'True
            Caption         =   "Temps pour réalisée la Quete :"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   60
            Top             =   360
            Width           =   2160
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Réponse à la fin de la quête :"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   4920
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nom de la quête :"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   510
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description de la quête :"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   1725
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Type de la quête :"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1260
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmEditeurQuetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Public Init As Boolean

Private Sub cases_Change()
frmEditeurQuetes.Label26.Caption = "Case scripter à éxécuter : " & frmEditeurQuetes.cases.Text
If Val(cases) < 0 Then MsgBox "Entrez un chiffre supérieur à zéro s'il vous plait."
End Sub

Private Sub collco_Click()
carted.Text = CoordM
xd.Text = CoordX
yd.Text = CoordY
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call LoadQuete(EditorIndex)
frmIndex.SetFocus
End Sub

Private Sub anul_Click()
Call LoadQuete(EditorIndex)
Unload Me
frmIndex.lstIndex.Clear
' Add the names
For i = 1 To MAX_QUETES
    frmIndex.lstIndex.AddItem i & " : " & Trim$(quete(i).nom)
Next i
frmIndex.SetFocus
End Sub

Private Sub defxy_Click()
frmEditeurQuetes.Hide
frmMirage.Show
frmMirage.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Long
Init = False
tpe.ListIndex = 0
For i = 1 To 6
    frtp(i).Visible = False
Next i

Label4.Caption = "Index de l'objet (pour la quête) : " & indo.value
Label15.Caption = "Index du PNJ (pour la quête) : " & indpnj.value
Label12.Caption = "Nombre de fois qu'il faut le tuer : " & nbt.value
Label22.Caption = "Nombre de points d'experiences a gagner : " & nbxp.value
Label11.Caption = "Numéro du PNJ : " & numepnj.value & " : " & Npc(numepnj.value).name
Label6.Caption = "Numéro de l'objet : " & numo.value & " : " & Item(numo.value).name
Label8.Caption = "Numéro de l'objet donné : " & numod.value & " : " & Item(numod.value).name
Label13.Caption = "Numéro du PNJ : " & numopnj.value & " : " & Npc(numopnj.value).name
Label9.Caption = "Numéro du PNJ : " & numpnj.value & " : " & Npc(numpnj.value).name
Label7.Caption = "Quantités à ramasser : " & quant.value
Label16.Caption = "Numéro de l'objet1 : " & ro1.value & " : " & Item(ro1.value).name
Label18.Caption = "Numéro de l'objet2 : " & ro2.value & " : " & Item(ro2.value).name
Label23.Caption = "Numéro de l'objet3 : " & ro3.value & " : " & Item(ro3.value).name
Label17.Caption = "Quantités de l'objet1 : " & rq1.value
Label21.Caption = "Quantités de l'objet2 : " & rq2.value
Label24.Caption = "Quantités de l'objet3 : " & rq3.value
Label14.Caption = "Points d'expérience gagnés : "
frmEditeurQuetes.Label26.Caption = "Case scripter a éxécuter : " & frmEditeurQuetes.cases.Text

For i = 1 To 6
    If tempr(i).value > 0 Then tp(i).Caption = "Temps pour réaliser la quête : " & tempr(i).value & "s (" & Int(tempr(i).value / 60) & "min" & tempr(i).value - (Int(tempr(i).value / 60) * 60) & "s)" Else tp(i).Caption = "Temps pour réaliser la quête : Infini"
Next i

numod.Max = MAX_ITEMS
numo.Max = MAX_ITEMS
ro1.Max = MAX_ITEMS
ro2.Max = MAX_ITEMS
ro3.Max = MAX_ITEMS
numpnj.Max = MAX_NPCS
numopnj.Max = MAX_NPCS
numepnj.Max = MAX_NPCS
End Sub

Private Sub indo_Change()
Label4.Caption = "Index de l'objet (pour la quête) : " & indo.value
If quete(EditorIndex).indexe(indo.value).Data1 > 0 Then numo.value = quete(EditorIndex).indexe(indo.value).Data1 Else numo.value = 1
If quete(EditorIndex).indexe(indo.value).Data2 > 0 Then quant.value = quete(EditorIndex).indexe(indo.value).Data2 Else quant.value = 0
End Sub

Private Sub indpnj_Change()
Label15.Caption = "Index du PNJ (pour la quête) : " & indpnj.value
If quete(EditorIndex).indexe(indpnj.value).Data1 > 0 Then numopnj.value = quete(EditorIndex).indexe(indpnj.value).Data1 Else numopnj.value = 1
If quete(EditorIndex).indexe(indpnj.value).Data2 > 0 Then nbt.value = quete(EditorIndex).indexe(indpnj.value).Data2 Else nbt.value = 0
End Sub

Private Sub nbt_Change()
Label12.Caption = "Nombre de fois qu'il faut le tuer : " & nbt.value
quete(EditorIndex).indexe(indpnj.value).Data2 = nbt.value
End Sub

Private Sub nbxp_Change()
Label22.Caption = "Nombre de points d'experiences a gagner : " & nbxp.value
End Sub

Private Sub numepnj_Change()
Label11.Caption = "Numéro du PNJ : " & numepnj.value & " : " & Npc(numepnj.value).name
End Sub

Private Sub numo_Change()
Label6.Caption = "Numéro de l'objet : " & numo.value & " : " & Item(numo.value).name
quete(EditorIndex).indexe(indo.value).Data1 = numo.value
End Sub

Private Sub numod_Change()
Label8.Caption = "Numéro de l'objet donné : " & numod.value & " : " & Item(numod.value).name
End Sub

Private Sub numopnj_Change()
Label13.Caption = "Numéro du PNJ : " & numopnj.value & " : " & Npc(numopnj.value).name
quete(EditorIndex).indexe(indpnj.value).Data1 = numopnj.value
End Sub

Private Sub numpnj_Change()
Label9.Caption = "Numéro du PNJ : " & numpnj.value & " : " & Npc(numpnj.value).name
End Sub

Private Sub OK_Click()

If Val(xd.Text) > 30 Then Call MsgBox("Veuillez mettre des coordonnées en X en chiffre et inférieur à 30"): Exit Sub
If Val(yd.Text) > 30 Then Call MsgBox("Veuillez mettre des coordonnées en Y en chiffre et inférieur à 30"): Exit Sub

quete(EditorIndex).nom = nom.Text
quete(EditorIndex).description = description.Text
quete(EditorIndex).reponse = reponse.Text

If tpe.ListIndex > 0 Then quete(EditorIndex).Temps = tempr(tpe.ListIndex).value

quete(EditorIndex).Type = tpe.ListIndex

If tpe.ListIndex = QUETE_TYPE_RECUP Then
    quete(EditorIndex).Data1 = indo.value
    quete(EditorIndex).Data2 = numo.value
    quete(EditorIndex).Data3 = quant.value
    quete(EditorIndex).String1 = vbNullString
ElseIf tpe.ListIndex = QUETE_TYPE_APORT Then
    quete(EditorIndex).Data1 = numod.value
    quete(EditorIndex).Data2 = numpnj.value
    quete(EditorIndex).Data3 = 0
    quete(EditorIndex).String1 = reppnj.Text
    For i = 1 To 15
        quete(EditorIndex).indexe(i).Data1 = 1
        quete(EditorIndex).indexe(i).Data2 = 0
        quete(EditorIndex).indexe(i).Data3 = 0
        quete(EditorIndex).indexe(i).String1 = vbNullString
    Next i
ElseIf tpe.ListIndex = QUETE_TYPE_PARLER Then
    quete(EditorIndex).Data1 = numepnj.value
    quete(EditorIndex).Data2 = 0
    quete(EditorIndex).Data3 = 0
    quete(EditorIndex).String1 = vbNullString
    For i = 1 To 15
        quete(EditorIndex).indexe(i).Data1 = 1
        quete(EditorIndex).indexe(i).Data2 = 0
        quete(EditorIndex).indexe(i).Data3 = 0
        quete(EditorIndex).indexe(i).String1 = vbNullString
    Next i
ElseIf tpe.ListIndex = QUETE_TYPE_TUER Then
    quete(EditorIndex).Data1 = indpnj.value
    quete(EditorIndex).Data2 = numopnj.value
    quete(EditorIndex).Data3 = nbt.value
    quete(EditorIndex).String1 = vbNullString
ElseIf tpe.ListIndex = QUETE_TYPE_FINIR Then
    quete(EditorIndex).Data1 = Val(xd.Text)
    quete(EditorIndex).Data2 = Val(yd.Text)
    quete(EditorIndex).Data3 = Val(carted.Text)
    quete(EditorIndex).String1 = vbNullString
    For i = 1 To 15
        quete(EditorIndex).indexe(i).Data1 = 1
        quete(EditorIndex).indexe(i).Data2 = 0
        quete(EditorIndex).indexe(i).Data3 = 0
        quete(EditorIndex).indexe(i).String1 = vbNullString
    Next i
ElseIf tpe.ListIndex = QUETE_TYPE_GAGNE_XP Then
    quete(EditorIndex).Data1 = nbxp.value
    quete(EditorIndex).Data2 = 0
    quete(EditorIndex).Data3 = 0
    quete(EditorIndex).String1 = vbNullString
    For i = 1 To 15
        quete(EditorIndex).indexe(i).Data1 = 1
        quete(EditorIndex).indexe(i).Data2 = 0
        quete(EditorIndex).indexe(i).Data3 = 0
        quete(EditorIndex).indexe(i).String1 = vbNullString
    Next i
End If

'If tpe.ListIndex = QUETE_TYPE_SCRIPT Then
'quete(EditorIndex).Data1 = numcase.value
'quete(EditorIndex).Data2 = 0
'quete(EditorIndex).Data3 = 0
'quete(EditorIndex).String1 = vbNullString
'For i = 1 To 15
'quete(EditorIndex).indexe(i).Data1 = 1
'quete(EditorIndex).indexe(i).Data2 = 0
'quete(EditorIndex).indexe(i).Data3 = 0
'quete(EditorIndex).indexe(i).String1 = vbNullString
'Next i
'End If

quete(EditorIndex).Recompence.exp = Val(rexp.Text)
quete(EditorIndex).Recompence.objn1 = ro1.value
quete(EditorIndex).Recompence.objn2 = ro2.value
quete(EditorIndex).Recompence.objn3 = ro3.value
quete(EditorIndex).Recompence.objq1 = rq1.value
quete(EditorIndex).Recompence.objq2 = rq2.value
quete(EditorIndex).Recompence.objq3 = rq3.value
quete(EditorIndex).Case = Val(cases.Text)

Call SendSaveQuete(EditorIndex)
Unload Me
frmIndex.lstIndex.Clear
' Add the names
For i = 1 To MAX_QUETES
    frmIndex.lstIndex.AddItem i & " : " & Trim$(quete(i).nom)
Next i
frmIndex.SetFocus
End Sub

Private Sub quant_Change()
Label7.Caption = "Quantités à ramasser : " & quant.value
quete(EditorIndex).indexe(indo.value).Data2 = quant.value
End Sub

Private Sub rexp_Change()
On Error Resume Next
If Not IsNumeric(rexp.Text) Then Call MsgBox("Les points d'expériences doivent être en chiffre."): Call rexp.SetFocus
End Sub

Private Sub rexp_LostFocus()
On Error Resume Next
If Not IsNumeric(rexp.Text) Then Call MsgBox("Les points d'expériences doivent être en chiffre."): Call rexp.SetFocus
End Sub

Private Sub ro1_Change()
Label16.Caption = "Numéro de l'objet1 : " & ro1.value & " : " & Item(ro1.value).name
End Sub

Private Sub ro2_Change()
Label18.Caption = "Numéro de l'objet2 : " & ro2.value & " : " & Item(ro2.value).name
End Sub

Private Sub ro3_Change()
Label23.Caption = "Numéro de l'objet3 : " & ro3.value & " : " & Item(ro3.value).name
End Sub

Private Sub rq1_Change()
Label17.Caption = "Quantités de l'objet1 : " & rq1.value
End Sub

Private Sub rq2_Change()
Label21.Caption = "Quantités de l'objet2 : " & rq2.value
End Sub

Private Sub rq3_Change()
Label24.Caption = "Quantités de l'objet3 : " & rq3.value
End Sub

Private Sub tempr_Change(Index As Integer)
If tempr(Index).value > 0 Then tp(Index).Caption = "Temps pour réaliser la quête : " & tempr(Index).value & "s (" & Int(tempr(Index).value / 60) & "min" & tempr(Index).value - (Int(tempr(Index).value / 60) * 60) & "s)" Else tp(Index).Caption = "Temps pour réaliser la quête : Infini"
End Sub

Private Sub tpe_Click()
If Init Then Call NetQueteType(EditorIndex)
For i = 1 To 6
    frtp(i).Visible = False
Next i
If tpe.ListIndex > 0 Then frtp(tpe.ListIndex).Visible = True
End Sub
