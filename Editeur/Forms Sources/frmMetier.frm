VERSION 5.00
Begin VB.Form frmMetier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Metier"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frChasseur 
      Caption         =   "Chasseur"
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4335
      Begin VB.HScrollBar scrlCibleNPC 
         Height          =   255
         Left            =   120
         Max             =   9
         TabIndex        =   16
         Top             =   480
         Width           =   4095
      End
      Begin VB.HScrollBar scrlExpNPC 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   13
         Top             =   1560
         Value           =   1
         Width           =   4095
      End
      Begin VB.HScrollBar scrlNPCNum 
         Height          =   255
         Left            =   120
         Max             =   200
         TabIndex        =   10
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblCibleNPC 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Cible"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblExpNPC 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Expérience rapportée"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblNPCNum 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Numéro NPC"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4335
   End
   Begin VB.ComboBox CMetier 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmMetier.frx":0000
      Left            =   120
      List            =   "frmMetier.frx":000A
      TabIndex        =   5
      Text            =   "Tuer un/des PNJ(s)"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TxtNom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Type de Métier"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Nom du métier:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMetier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call MetierEditorCancel
End Sub

Private Sub cmdOk_Click()
    Call MetierEditorOk
End Sub

Private Sub CMetier_Change()
    If CMetier.ListIndex = 0 Then
        frChasseur.Caption = "Tuer"
        Label4.Caption = "Numéro PNJ"
    ElseIf CMetier.ListIndex = 1 Then
        frChasseur.Caption = "Craft"
        Label4.Caption = "Numéro Recette"
    End If
End Sub

Private Sub Form_Load()
    scrlNPCNum.Max = MAX_NPCS
    scrlCibleNPC.Max = MAX_DATA_METIER
End Sub

Private Sub scrlCibleNPC_Change()
    lblCibleNPC.Caption = scrlCibleNPC.value
    If Metier(EditorIndex).Data(scrlCibleNPC.value, 1) > 0 Then scrlExpNPC.value = Metier(EditorIndex).Data(scrlCibleNPC.value, 1)
    scrlNPCNum.value = Metier(EditorIndex).Data(scrlCibleNPC.value, 0)
End Sub

Private Sub scrlExpNPC_Change()
    lblExpNPC.Caption = scrlExpNPC.value
    Metier(EditorIndex).Data(scrlCibleNPC.value, 1) = scrlExpNPC.value
End Sub

Private Sub scrlNPCNum_Change()
    If CMetier.ListIndex = 0 Then
        If scrlNPCNum.value > 0 Then
            lblNPCNum.Caption = scrlNPCNum.value & ": " & Npc(scrlNPCNum.value).name
        Else
            lblNPCNum.Caption = "Pas de PNJ"
        End If
    ElseIf CMetier.ListIndex = 1 Then
        If scrlNPCNum.value > 0 Then
            lblNPCNum.Caption = scrlNPCNum.value & ": " & recette(scrlNPCNum.value).nom
        Else
            lblNPCNum.Caption = "Pas de Craft"
        End If
    End If
    Metier(EditorIndex).Data(scrlCibleNPC.value, 0) = scrlNPCNum.value
End Sub
