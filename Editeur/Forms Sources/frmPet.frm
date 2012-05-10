VERSION 5.00
Begin VB.Form frmPets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Familier"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   LinkTopic       =   "Editeur de Famillier"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   199
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   277
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScrl 
      Height          =   1095
      Left            =   3840
      TabIndex        =   16
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar HScrl 
      Height          =   255
      Left            =   2760
      TabIndex        =   15
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   13
      Top             =   720
      Width           =   1095
      Begin VB.PictureBox PictApp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.HScrollBar ScrlDefence 
      Height          =   255
      Left            =   180
      Max             =   255
      TabIndex        =   11
      Top             =   2520
      Width           =   1935
   End
   Begin VB.HScrollBar ScrlForce 
      Height          =   255
      Left            =   180
      Max             =   255
      TabIndex        =   8
      Top             =   1800
      Width           =   1935
   End
   Begin VB.HScrollBar ScrlApp 
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1140
      Width           =   1935
   End
   Begin VB.TextBox TxtNom 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblDefence 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Défense:"
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   2160
      Width           =   2115
   End
   Begin VB.Label lblForce 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Force:"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1500
      Width           =   2055
   End
   Begin VB.Label lblAppNum 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Apparence :"
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Nom :"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call PetEditorCancel
End Sub

Private Sub cmdOk_Click()
    Call PetEditorOk
End Sub

Private Sub Form_Load()
    ScrlApp.Max = MAX_DX_PETS
End Sub

Private Sub HScrl_Change()
    PictApp.Left = HScrl.value * -1
End Sub

Private Sub ScrlApp_Change()
    lblAppNum.Caption = ScrlApp.value
    PreparePet (ScrlApp.value)
    frmPets.PictApp.Picture = LoadPNG(App.Path & "\GFX\Pets\Pet" & ScrlApp.value & ".png")
    PictApp.Width = DDSD_Pets(ScrlApp.value).lWidth / 4
    PictApp.Height = DDSD_Pets(ScrlApp.value).lHeight / 4
End Sub

Private Sub ScrlDefence_Change()
    lblDefence.Caption = ScrlDefence.value
End Sub

Private Sub ScrlForce_Change()
    lblForce.Caption = ScrlForce.value
End Sub

Private Sub VScrl_Change()
    PictApp.Top = VScrl.value * -1
End Sub
