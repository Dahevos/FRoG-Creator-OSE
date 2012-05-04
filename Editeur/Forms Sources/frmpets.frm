VERSION 5.00
Begin VB.Form frmpet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crédits"
   ClientHeight    =   3960
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   4905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "Revenir au menu"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label creditline1 
      BackStyle       =   0  'Transparent
      Caption         =   "Merci à Hinomi pour sa belle bannière et à Rose pour la partie graphique de FRoG Creator"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmpets.frx":0000
      Height          =   1395
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remerciements : Coke, GodSentdeath, Katsuo, Edouard, Dahevos et à toute la communauté de FRoG Creator."
      Height          =   795
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmpet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
If frmMirage.Visible Then
    frmpet.Visible = False
    frmMirage.SetFocus
Else
    frmMainMenu.Visible = True
    frmpet.Visible = False
End If
End Sub

