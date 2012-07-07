VERSION 5.00
Begin VB.Form frmMapErr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRoG Creator - Erreur lors de la fermeture"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNormal 
      Caption         =   "Démarrer normalement"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdRestaure 
      Caption         =   "Restaurer la carte précédente"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Informations sur la sauvegarde : Carte numéros 1, Dernière modification : 10/09/09 15h30"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmMapErr.frx":0000
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmMapErr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private curMap As Long

Private Sub cmdNormal_Click()
    Call Unload(Me)
End Sub

Private Sub cmdRestaure_Click()
    If curMap > 0 Then
        'Supprimer l'ancien fichier
        If FileExists(App.Path & "\Maps\map" & curMap & ".fcc") Then Kill App.Path & "\Maps\map" & curMap & ".fcc"
        'Copier le nouveau fichier
        Call FileCopy(App.Path & "\Maps\map" & curMap & "BACKUP.fcc", App.Path & "\Maps\map" & curMap & ".fcc")
        'Fermer la fenêtre
        Call Unload(Me)
    End If
End Sub

Private Sub Form_Load()
    curMap = 0
End Sub

Public Sub Init(ByVal Map As Long)
On Error Resume Next
Dim Dmod As String
    'Initialisation de la map courante et de la fenêtre si le fichier de backup existe
    If FileExists(App.Path & "\Maps\map" & Map & "BACKUP.fcc") Then
        Dmod = FileDateTime(App.Path & "\Maps\map" & Map & "BACKUP.fcc")
        Label2.Caption = "Informations sur la sauvegarde : Carte " & Map & ", Derniére modification : " & Val(Day(Dmod)) & "/" & Val(Month(Dmod)) & "/" & Val(Year(Dmod)) & " " & Val(Hour(Dmod)) & "h" & Val(minute(Dmod))
        curMap = Map
        Me.Show
    End If
End Sub
