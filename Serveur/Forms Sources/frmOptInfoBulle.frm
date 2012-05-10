VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptInfoBulle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options des Infos Bulles"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd 
      Left            =   4320
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton sauv 
      Caption         =   "Sauvegarder"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Messages à afficher"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox acoul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4080
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1320
         Width           =   495
      End
      Begin VB.PictureBox jcoul 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4080
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   480
         Width           =   495
      End
      Begin VB.CheckBox ma 
         Caption         =   "Messages sur les admins"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.CheckBox mer 
         Caption         =   "Messages sur les erreurs"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox mj 
         Caption         =   "Messages sur les joueurs"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmOptInfoBulle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ma_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ma.value = Checked Then IBAdmin = False Else IBAdmin = True
End Sub

Private Sub mer_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If mer.value = Checked Then IBErr = False Else IBErr = True
End Sub

Private Sub mj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If mj.value = Checked Then IBJoueur = False Else IBJoueur = True
End Sub
Private Sub form_load()
If IBJoueur Then mj.value = 1
If IBErr Then mer.value = 1
If IBAdmin Then ma.value = 1
End Sub

Private Sub sauv_Click()
Call SauvIBOpt
Unload Me
End Sub
