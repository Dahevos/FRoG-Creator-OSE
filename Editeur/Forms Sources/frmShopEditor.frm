VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmShopEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer un magasin"
   ClientHeight    =   8835
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   11790
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
   ScaleHeight     =   8835
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   360
      TabIndex        =   19
      ToolTipText     =   "Représentation schématique du magasin et de ses différentes parties"
      Top             =   4920
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   353
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Arme"
      TabPicture(0)   =   "frmShopEditor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstTradeItem(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Bouclier"
      TabPicture(1)   =   "frmShopEditor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstTradeItem(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Armure"
      TabPicture(2)   =   "frmShopEditor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstTradeItem(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Casque"
      TabPicture(3)   =   "frmShopEditor.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstTradeItem(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Sort"
      TabPicture(4)   =   "frmShopEditor.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstTradeItem(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Divers"
      TabPicture(5)   =   "frmShopEditor.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lstTradeItem(5)"
      Tab(5).ControlCount=   1
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   5
         ItemData        =   "frmShopEditor.frx":00A8
         Left            =   -74880
         List            =   "frmShopEditor.frx":00AA
         TabIndex        =   25
         Top             =   360
         Width           =   10815
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   4
         ItemData        =   "frmShopEditor.frx":00AC
         Left            =   -74880
         List            =   "frmShopEditor.frx":00AE
         TabIndex        =   24
         Top             =   360
         Width           =   10815
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   3
         ItemData        =   "frmShopEditor.frx":00B0
         Left            =   -74880
         List            =   "frmShopEditor.frx":00B2
         TabIndex        =   23
         Top             =   360
         Width           =   10815
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   2
         ItemData        =   "frmShopEditor.frx":00B4
         Left            =   -74880
         List            =   "frmShopEditor.frx":00B6
         TabIndex        =   22
         Top             =   360
         Width           =   10815
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   1
         ItemData        =   "frmShopEditor.frx":00B8
         Left            =   -74880
         List            =   "frmShopEditor.frx":00BA
         TabIndex        =   21
         Top             =   360
         Width           =   10815
      End
      Begin VB.ListBox lstTradeItem 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Index           =   0
         ItemData        =   "frmShopEditor.frx":00BC
         Left            =   120
         List            =   "frmShopEditor.frx":00BE
         TabIndex        =   20
         Top             =   360
         Width           =   10815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Propriétés du magasin"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   11535
      Begin VB.ComboBox cmbItemGive 
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
         ItemData        =   "frmShopEditor.frx":00C0
         Left            =   1200
         List            =   "frmShopEditor.frx":00C2
         Style           =   2  'Dropdown List
         TabIndex        =   14
         ToolTipText     =   "Objet donné aux joueurs en échange de l'objet reçu"
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtItemGiveValue 
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
         Left            =   1200
         TabIndex        =   13
         Text            =   "1"
         ToolTipText     =   "Nombre d'objet(s) donné"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbItemGet 
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
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Objet prit aux joueurs en échange de l'objet donné"
         Top             =   360
         Width           =   3975
      End
      Begin VB.TextBox txtItemGetValue 
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
         Left            =   6480
         TabIndex        =   11
         Text            =   "1"
         ToolTipText     =   "Nombre d'objet(s) reçu"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Mettre dans le slot sélectionné"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         ToolTipText     =   "Mettre les objets dans le magasin dans le slot sélectionner"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
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
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Objet reçu :"
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
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Left            =   5640
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
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
      Left            =   3960
      TabIndex        =   3
      ToolTipText     =   "Quitte la fenêtre d'édition et enregistre le magasin"
      Top             =   8280
      Width           =   1575
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
      TabIndex        =   4
      ToolTipText     =   "Quitte la fenêtre d'édition sans enregistrer le magasin"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtLeaveSay 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Message que verront les joueurs quand ils partiront du magasin"
      Top             =   1680
      Width           =   3975
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
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Nom du magasin"
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtJoinSay 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Message que verront les joueurs quand ils entreront dans le magasin"
      Top             =   1680
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Propriétés Générale"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   11535
      Begin VB.ComboBox cmbItemFix 
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
         ItemData        =   "frmShopEditor.frx":00C4
         Left            =   5040
         List            =   "frmShopEditor.frx":00C6
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Objet donné aux joueurs en échange de l'objet reçu"
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CheckBox chkFixesItems 
         Caption         =   "Le magasin répare les objets"
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
         Left            =   5880
         TabIndex        =   26
         ToolTipText     =   "Si cette casse est cocher le magasin pourras réparer les objet user"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Objet pour payer les réparations :"
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
         Left            =   2760
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Message d'accueil :"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1185
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
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Message au départ :"
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
         Left            =   5880
         TabIndex        =   6
         Top             =   1320
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmShopEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFixesItems_Click()
If chkFixesItems.value = Checked Then cmbItemFix.Visible = True: Label7.Visible = True Else cmbItemFix.Visible = False: Label7.Visible = False
End Sub

Private Sub cmdOk_Click()
    Call ShopEditorOk
End Sub

Private Sub cmdCancel_Click()
    Call ShopEditorCancel
End Sub

Private Sub cmdUpdate_Click()
Dim Index As Long, i As Long, ItemNum As Long
    
    Index = lstTradeItem(SSTab1.Tab).ListIndex + 1
    i = SSTab1.Tab + 1
    ItemNum = cmbItemGet.ListIndex
    
    If ItemNum > 0 Then
        If i = 1 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Cliquer sur la rubrique BOUCLIER pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Cliquer sur la rubrique ARMURE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Cliquer sur la rubrique CASQUE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Cliquer sur la rubrique SORT pour ajouter ceci."
                Exit Sub
            Else
                MsgBox "Cliquer sur la rubrique DIVERS pour ajouter ceci."
                Exit Sub
            End If
        ElseIf i = 2 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Cliquer sur la rubrique ARME pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Cliquer sur la rubrique ARMURE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Cliquer sur la rubrique CASQUE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Cliquer sur la rubrique SORT pour ajouter ceci."
                Exit Sub
            Else
                MsgBox "Cliquer sur la rubrique DIVERS pour ajouter ceci."
                Exit Sub
            End If
        ElseIf i = 3 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Cliquer sur la rubrique ARME pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Cliquer sur la rubrique BOUCLIER pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Cliquer sur la rubrique CASQUE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Cliquer sur la rubrique SORT pour ajouter ceci."
                Exit Sub
            Else
                MsgBox "Cliquer sur la rubrique DIVERS pour ajouter ceci."
                Exit Sub
            End If
        ElseIf i = 4 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Cliquer sur la rubrique ARME pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Cliquer sur la rubrique BOUCLIER pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Cliquer sur la rubrique ARMURE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Cliquer sur la rubrique SORT pour ajouter ceci."
                Exit Sub
            Else
                MsgBox "Cliquer sur la rubrique DIVERS pour ajouter ceci."
                Exit Sub
            End If
        ElseIf i = 5 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Cliquer sur la rubrique ARME pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Cliquer sur la rubrique BOUCLIER pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Cliquer sur la rubrique ARMURE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Cliquer sur la rubrique CASQUE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
            Else
                MsgBox "Cliquer sur la rubrique SORT pour ajouter ceci."
                Exit Sub
            End If
        ElseIf i = 6 Then
            If Item(ItemNum).Type = ITEM_TYPE_WEAPON Then
                MsgBox "Cliquer sur la rubrique ARME pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SHIELD Then
                MsgBox "Cliquer sur la rubrique BOUCLIER pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_ARMOR Then
                MsgBox "Cliquer sur la rubrique ARMURE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_HELMET Then
                MsgBox "Cliquer sur la rubrique CASQUE pour ajouter ceci."
                Exit Sub
            ElseIf Item(ItemNum).Type = ITEM_TYPE_SPELL Then
                MsgBox "Cliquer sur la rubrique DIVERS pour ajouter ceci."
                Exit Sub
            Else
            End If
        End If
    End If
    
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).value(Index).GiveItem = cmbItemGive.ListIndex
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).value(Index).GiveValue = Val(txtItemGiveValue.Text)
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).value(Index).GetItem = cmbItemGet.ListIndex
    Shop(EditorIndex).TradeItem(SSTab1.Tab + 1).value(Index).GetValue = Val(txtItemGetValue.Text)

    Call UpdateShopTrade
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call ShopEditorCancel
End Sub
