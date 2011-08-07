VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mise à jours automatique de l'Editeur Externe de Frog Creator"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6300
   ControlBox      =   0   'False
   Icon            =   "frmUpdates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   420
   StartUpPosition =   2  'CenterScreen
   Begin FrogCreator.Download DL 
      Height          =   1425
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   2514
   End
   Begin VB.TextBox news 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmUpdates.frx":08CA
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton label1 
      Caption         =   "Fermer"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label L3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5895
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public exe As String

Private Sub BarColor1_Change()

End Sub

Private Sub DL_BeginDownload(URL As String)
FileBool = False
Label4.Caption = "0 %"
End Sub

Private Sub DL_Error(SubName As String, ErrNum As Long, ErrDesc As String)
Select Case ErrNum
Case 53: Call MsgBox("Vérifiez avec l'administrateur si votre fichier de configuration est à jour", vbCritical, "Erreur")
Case -2147012889: Call MsgBox("Vérifiez que votre connexion est active ou si votre parefeu/routeur bloque l'Accès à l'updateur", vbCritical, "Erreur")
Case Else: Call MsgBox("Une erreur est survenu dans l'OCXUpdate, Veuillez contacter l'administrateur si le problême persiste" & vbNewLine & vbNewLine & "Err #" & ErrNum & vbNewLine & ErrDesc, vbCritical, "Erreur")
End Select
DL.EndDownload
UpdateErr = True
End Sub

Private Sub DL_FinishDownload(File As String, FileStr As String)
FileBool = True
End Sub

Private Sub DL_OnProgress(BytesRead As Long, BytesMax As Long, ProgressPercent As Byte)
Label4.Caption = ProgressPercent & " %   -   " & Round(BytesRead / 1024, 2) & "Ko / " & Round(BytesMax / 1024, 2) & "Ko"
End Sub

Private Sub DL_Unload()
Call GameDestroy
End Sub

Private Sub Label1_Click()
'If Not Uexe Then
'    MsgBox "L'éditeur doit ce relancer pour prendre en compte la mise a jour. Merci de le relancer."
'    Unload Me
'    Call GameDestroy
'Else
'    Unload Me
'    Call Shell(App.Path & "\" & Uexen, vbNormalFocus)
'    Call GameDestroy
'End If
MsgBox "L'éditeur doit se relancer pour prendre en compte la mise a jour."
DL.EndDownload
Call GameDestroy
End Sub

