VERSION 5.00
Begin VB.Form frmKeepNotes 
   BorderStyle     =   0  'None
   Caption         =   "Note du joueur"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   0
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmKeepNotes.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Close 
      Appearance      =   0  'Flat
      Caption         =   "Fermer"
      Height          =   300
      Left            =   2880
      TabIndex        =   2
      Top             =   5400
      Width           =   2070
   End
   Begin VB.CommandButton Save 
      Appearance      =   0  'Flat
      BackColor       =   &H00789298&
      Caption         =   "Sauvegarder"
      Height          =   300
      Left            =   285
      MaskColor       =   &H00789298&
      TabIndex        =   1
      Top             =   5400
      Width           =   2070
   End
   Begin VB.TextBox Notetext 
      Appearance      =   0  'Flat
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4680
   End
End
Attribute VB_Name = "frmKeepNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Close_Click()
frmKeepNotes.Visible = False
End Sub



Private Sub Save_Click()
Dim iFileNum As Integer

'Get a free file handle
iFileNum = FreeFile

'If the file is not there, one will be created
'If the file does exist, this one will
'overwrite it.
Open App.Path & "\notes.txt" For Output As iFileNum

Print #iFileNum, Notetext.Text

Close iFileNum

End Sub

