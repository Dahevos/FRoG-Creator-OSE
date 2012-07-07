VERSION 5.00
Begin VB.Form frmChars 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sélection du personnage"
   ClientHeight    =   3210
   ClientLeft      =   105
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChars.frx":0000
   ScaleHeight     =   3210
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton picCancel 
      Caption         =   "Retour"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton picUseChar 
      Caption         =   "Utiliser ce personnage"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ListBox lstChars 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1680
      ItemData        =   "frmChars.frx":54AF2
      Left            =   480
      List            =   "frmChars.frx":54AF4
      TabIndex        =   0
      Top             =   120
      Width           =   3705
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picUseChar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMirage.Timer2.Enabled = True
End Sub

Private Sub lstChars_DblClick()
Call picUseChar_Click
End Sub

Private Sub lstChars_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picUseChar_Click
End Sub

Private Sub picCancel_Click()
Dim i As Long
    For i = 1 To MAX_INV - 1
        Unload frmMirage.picInv(i)
    Next
    Call TcpDestroy
    frmMainMenu.Visible = True
    Me.Visible = False
    Sleep 2000
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then
        MsgBox "Il n'y a pas de personnage à cet emplacement."
        Exit Sub
    End If
    Call MenuState(MENU_STATE_USECHAR)
        
End Sub
