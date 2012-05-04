VERSION 5.00
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRoG Server 0.6.2"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   ControlBox      =   0   'False
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin Serveur.ctlProgressBar chrg 
      Height          =   375
      Left            =   0
      Top             =   360
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Initialisation ..."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   80
      Width           =   5535
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    NewDoEvents
    If InDestroy = False Then Call InitServer
End Sub

Private Sub Timer1_Timer()
If InDestroy = True Then Unload Me
End Sub
