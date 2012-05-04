VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Begin VB.Form frmSpriteChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attribut de changement de Sprite"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   3413
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   2822
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Sélection du Sprite"
      TabPicture(0)   =   "frmSpriteChange.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblSprite"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCost"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblItem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "scrlSprite"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdOk"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "scrlCost"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "scrlItem"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "picSprite"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCancel"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
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
         Left            =   3720
         TabIndex        =   2
         Top             =   1560
         Width           =   1035
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   3720
         ScaleHeight     =   68
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   68
         TabIndex        =   12
         Top             =   240
         Width           =   1020
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         LargeChange     =   10
         Left            =   180
         Max             =   30
         TabIndex        =   9
         Top             =   960
         Width           =   3495
      End
      Begin VB.HScrollBar scrlCost 
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         Max             =   30000
         TabIndex        =   4
         Top             =   1500
         Width           =   3495
      End
      Begin VB.CommandButton cmdOk 
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
         Left            =   3720
         TabIndex        =   3
         Top             =   1320
         Width           =   1035
      End
      Begin VB.HScrollBar scrlSprite 
         Height          =   255
         LargeChange     =   10
         Left            =   180
         Max             =   1000
         TabIndex        =   1
         Top             =   420
         Width           =   3435
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "Aucun Prix"
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
         Left            =   660
         TabIndex        =   11
         Top             =   780
         Width           =   660
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Objet:"
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
         TabIndex        =   10
         Top             =   780
         Width           =   405
      End
      Begin VB.Label lblCost 
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
         Left            =   660
         TabIndex        =   8
         Top             =   1320
         Width           =   1755
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valeur:"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sprite:"
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
         TabIndex        =   6
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblSprite 
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
         Left            =   660
         TabIndex        =   5
         Top             =   240
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmSpriteChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmSpriteChange.Visible = False
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub cmdOk_Click()
    SpritePic = scrlSprite.value
    SpriteItem = scrlItem.value
    SpritePrice = scrlCost.value
    scrlCost.value = 0
    scrlSprite.value = 0
    scrlItem.value = 0
    frmSpriteChange.Visible = False
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub Form_Load()
    scrlSprite.Max = MAX_DX_SPRITE
    If SpritePic < scrlSprite.Min Then SpritePic = scrlSprite.Min
    scrlSprite.value = SpritePic
    If SpriteItem < scrlItem.Min Then SpriteItem = scrlItem.Min
    scrlItem.value = SpriteItem
    If SpritePrice < scrlCost.Min Then SpritePrice = scrlCost.Min
    scrlCost.value = SpritePrice
    
    Call scrlSprite_Change
End Sub

Private Sub scrlCost_Change()
    lblCost.Caption = scrlCost.value
End Sub

Private Sub scrlItem_Change()
    If scrlItem.value = 0 Then lblItem.Caption = "Aucun Prix": Exit Sub Else lblItem.Caption = scrlItem.value & " - " & Trim$(Item(scrlItem.value).name)
    If Item(scrlItem.value).Type = ITEM_TYPE_CURRENCY Then scrlCost.Enabled = True Else scrlCost.Enabled = False
End Sub

Private Sub scrlSprite_Change()
On Error Resume Next
    lblSprite.Caption = scrlSprite.value
    Call PrepareSprite(scrlSprite.value)
    Call AffSurfPic(DD_SpriteSurf(scrlSprite.value), picSprite, 0, 0)
    'Call BitBlt(picSprite.hDC, 0, 0, PIC_X, PIC_Y * PIC_NPC1, picSprites.hDC, 3 * PIC_X, scrlSprite.value * (PIC_Y * PIC_NPC1), SRCCOPY)
End Sub
