VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptCoul 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options des couleurs"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   529
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton def 
      Caption         =   "Par défault"
      Height          =   255
      Left            =   4560
      TabIndex        =   33
      Top             =   4560
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   7080
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Couleurs des messages"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   31
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   3960
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   28
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   3600
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   27
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   3240
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   22
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   2880
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   21
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   2520
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   20
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   2160
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   19
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1800
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   14
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1440
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   13
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   12
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox MsgC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   11
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Message de guilde :"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3960
         Width           =   1425
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Message d'alerte :"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Message départ/arriver :"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3240
         Width           =   1740
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Message ""Qui est en ligne"" :"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   2025
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Message à distance :"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Message d'admin :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Message d'émotion :"
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hurlement (broadcast) :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1650
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Message d'aide :"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Message globale :"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Message de discussion :"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1740
      End
   End
   Begin VB.CommandButton sauv 
      Caption         =   "Sauvegarder"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Couleurs des accès"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox adm 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   4
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox dev 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   3
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox mapp 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   2
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1080
         Width           =   615
      End
      Begin VB.PictureBox modo 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   615
         TabIndex        =   1
         ToolTipText     =   "Cliquez pour modifier"
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Couleur de l'accès Admin :"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Couleur de l'accès Devellopeur :"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Couleur de l'accès Mapeur :"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1980
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Couleur de l'accès Modo :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1845
      End
   End
End
Attribute VB_Name = "frmOptCoul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub adm_Click()
cmd.Flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then adm.BackColor = cmd.Color
End Sub

Private Sub def_Click()
PutVar App.Path & "\Data.ini", "COULEURS", "AccAdmin", "16711935"
PutVar App.Path & "\Data.ini", "COULEURS", "AccDevelopeur", "8388608"
PutVar App.Path & "\Data.ini", "COULEURS", "AccModo", "8421504"
PutVar App.Path & "\Data.ini", "COULEURS", "AccMapeur", "8421376"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgDiscu", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgGlob", "32768"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgDist", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgHurl", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgEmot", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgAdmin", "16776960"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgAide", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgQui", "12632256"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgDep", "12632256"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgAlert", "16777215"
PutVar App.Path & "\Data.ini", "COULEURS", "MsgGuilde", "65280"
Call ChargOptCoul
End Sub

Private Sub dev_Click()
cmd.Flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then dev.BackColor = cmd.Color
End Sub


Private Sub mapp_Click()
cmd.Flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then mapp.BackColor = cmd.Color
End Sub

Private Sub modo_Click()
cmd.Flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then modo.BackColor = cmd.Color
End Sub

Private Sub MsgC_Click(Index As Integer)
cmd.Flags = &H2& + &H1&
cmd.ShowColor
If cmd.Color > -1 Then MsgC(Index).BackColor = cmd.Color
End Sub

Private Sub sauv_Click()
AccAdmin = adm.BackColor
AccDevelopeur = dev.BackColor
AccModo = modo.BackColor
AccMapeur = mapp.BackColor
Call PutVar(App.Path & "\Data.ini", "COULEURS", "AccAdmin", adm.BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "AccDevelopeur", dev.BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "AccModo", modo.BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "AccMapeur", mapp.BackColor)
Call SendDataToAll("PICVALUE" & SEP_CHAR & PIC_PL & SEP_CHAR & PIC_NPC1 & SEP_CHAR & PIC_NPC2 & SEP_CHAR & AccModo & SEP_CHAR & AccMapeur & SEP_CHAR & AccDevelopeur & SEP_CHAR & AccAdmin & END_CHAR)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgDiscu", MsgC(0).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgGlob", MsgC(1).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgDist", MsgC(2).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgHurl", MsgC(3).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgEmot", MsgC(4).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgAdmin", MsgC(5).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgAide", MsgC(6).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgQui", MsgC(7).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgDep", MsgC(8).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgAlert", MsgC(9).BackColor)
Call PutVar(App.Path & "\Data.ini", "COULEURS", "MsgGuilde", MsgC(10).BackColor)
Unload Me
End Sub
