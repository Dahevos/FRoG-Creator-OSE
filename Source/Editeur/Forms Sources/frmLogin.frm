VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Connexion"
   ClientHeight    =   5835
   ClientLeft      =   180
   ClientTop       =   405
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   5115
   Begin VB.Label picCancel 
      AutoSize        =   -1  'True
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   " Revenir"
      Height          =   195
      Left            =   2280
      TabIndex        =   0
      Top             =   5400
      Width           =   600
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmLogin.Visible = False
End Sub


