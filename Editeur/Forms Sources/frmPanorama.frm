VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPanorama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sélection d'un un panorama pour la carte"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   353
      TabMaxWidth     =   1764
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Panorama"
      TabPicture(0)   =   "frmPanorama.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstPano"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOk"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdCancel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   4320
         Width           =   1695
      End
      Begin VB.ListBox lstPano 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         ItemData        =   "frmPanorama.frx":001C
         Left            =   240
         List            =   "frmPanorama.frx":0023
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmPanorama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Val(lstPano.Tag) = 1 Then frmMapProperties.PanoSup.Text = lstPano.Text Else frmMapProperties.PanoInf.Text = lstPano.Text
If Val(lstPano.Tag) = 1 And lstPano.ListIndex = 0 Then frmMapProperties.PanoSup.Text = vbNullString
If Val(lstPano.Tag) = 0 And lstPano.ListIndex = 0 Then frmMapProperties.PanoInf.Text = vbNullString
If Len(lstPano.Text) > 50 Then Call MsgBox("Nom du fichier trop grand, veuillez en choisir un autre s'il vous plait", vbExclamation, "Erreur") Else Unload Me
End Sub


Private Sub lstPano_DblClick()
Call cmdOk_Click
End Sub
