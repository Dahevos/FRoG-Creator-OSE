VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3180
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000001&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   480
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyEscape) Then
        KeyAscii = 0
        Call DestroyDirectX
        Call StopMidi
        InGame = False
        frmMirage.Socket.Close
        frmMainMenu.Visible = True
        Connucted = False
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim R1 As Long, G1 As Long, B1 As Long
    Rep_Theme = ReadINI("Themes", "Theme", App.Path & "\Themes.ini")
    
    R1 = Val(ReadINI("BARE", "R", App.Path & Rep_Theme & "\Couleur.ini"))
    G1 = Val(ReadINI("BARE", "V", App.Path & Rep_Theme & "\Couleur.ini"))
    B1 = Val(ReadINI("BARE", "B", App.Path & Rep_Theme & "\Couleur.ini"))
    Shape1.FillColor = RGB(R1, G1, B1)
    
    R1 = Val(ReadINI("FOND", "R", App.Path & Rep_Theme & "\Couleur.ini"))
    G1 = Val(ReadINI("FOND", "V", App.Path & Rep_Theme & "\Couleur.ini"))
    B1 = Val(ReadINI("FOND", "B", App.Path & Rep_Theme & "\Couleur.ini"))
    Me.BackColor = RGB(R1, G1, B1)
End Sub

