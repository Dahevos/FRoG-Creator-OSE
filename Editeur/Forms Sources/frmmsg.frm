VERSION 5.00
Begin VB.Form frmmsg 
   BorderStyle     =   0  'None
   Caption         =   "Message"
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Erreur? cliquer sur le message pour Quitter!"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Veuillez patienter SVP ..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3075
   End
End
Attribute VB_Name = "frmmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
If GettingMap And FileExists(App.Path & "\Maps\" & Player(MyIndex).Map & ".fcc") Then Call Kill(App.Path & "\Maps\" & Player(MyIndex).Map & ".fcc")
Call GameDestroy
End Sub

Private Sub Label1_Click()
If GettingMap And FileExists(App.Path & "\Maps\" & Player(MyIndex).Map & ".fcc") Then Call Kill(App.Path & "\Maps\" & Player(MyIndex).Map & ".fcc")
Call GameDestroy
End Sub
