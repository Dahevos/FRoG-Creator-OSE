VERSION 5.00
Begin VB.Form frmcraft 
   BorderStyle     =   0  'None
   Caption         =   "Table de Craft"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmcraft.frx":0000
   ScaleHeight     =   5250
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCrafter 
      Caption         =   "Crafter !"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   3495
   End
   Begin VB.HScrollBar scrlRecettes 
      Height          =   220
      Left            =   120
      Max             =   9
      TabIndex        =   2
      Top             =   1100
      Width           =   4335
   End
   Begin VB.Label lblQuitter 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   16
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblQuitter 
      BackStyle       =   0  'Transparent
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   15
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblObtenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Obtenu"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblNeedItem 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label lblMetierNom 
      BackStyle       =   0  'Transparent
      Caption         =   "nom"
      Height          =   255
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label lblNom 
      BackStyle       =   0  'Transparent
      Caption         =   "Recette: "
      Height          =   255
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "frmcraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCrafter_Click()
    If RecetteSelect = Player(MyIndex).Metier Then
        If Metier(RecetteSelect).Data(scrlRecettes.value, 0) > 0 Then
            Call SendData("crafter" & SEP_CHAR & Metier(RecetteSelect).Data(scrlRecettes.value, 0) & END_CHAR)
        Else
            MsgBox ("Pas de recettes .")
        End If
    Else
        MsgBox ("Ce n'est pas votre métier .")
    End If
End Sub

Private Sub Form_Load()
    scrlRecettes.Max = MAX_DATA_METIER
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = True
drx = x
dry = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (x - drx), Me.Top + (y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.Height Then Me.Top = Screen.Height \ 2: Me.Left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub lblQuitter_Click(Index As Integer)
    Me.Hide
End Sub

Private Sub scrlRecettes_Change()
Dim i As Byte, n As Byte
    If Metier(RecetteSelect).Data(scrlRecettes.value, 0) > 0 Then
        lblNom.Caption = recette(Metier(RecetteSelect).Data(scrlRecettes.value, 0)).nom
        n = Metier(RecetteSelect).Data(scrlRecettes.value, 0)
        For i = 0 To 9
            If recette(n).InCraft(i, 0) > 0 Then
                lblNeedItem(i).Caption = Item(recette(n).InCraft(i, 0)).name & " (*" & recette(n).InCraft(i, 1) & ")"
            Else
                lblNeedItem(i).Caption = "Pas d'objet"
            End If
        Next i
        If recette(n).craft(0) > 0 Then
            lblObtenu.Caption = Item(recette(n).craft(0)).name & " (*" & recette(n).craft(1) & ")"
        Else
            lblObtenu.Caption = "Pas de craft"
        End If
    Else
        lblNom.Caption = "Pas de recettes"
        For i = 0 To 9
            lblNeedItem(i).Caption = "Pas d'objet"
        Next i
        lblObtenu.Caption = "Pas de craft"
    End If
End Sub
