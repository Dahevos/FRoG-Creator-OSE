VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmServerChooser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sélection du Serveur"
   ClientHeight    =   2610
   ClientLeft      =   2850
   ClientTop       =   1755
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmServerChooser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServerChooser.frx":000C
   ScaleHeight     =   2610
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstServers 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Quitter"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Connecter"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   480
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmServerChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GAME_IP As String, GAME_PORT As Long
Public Path As String
Public Extension As String
    Private Sub cmdCancel_Click()
        Call GameDestroy
        Unload Me
    End Sub

    Private Sub cmdOk_Click()
        If lstServers.ListCount <= 0 Then Exit Sub

        GAME_IP = ReadINI("SERVER" & lstServers.ListIndex, "IP", App.Path & "\Config\Serveur.ini")
        GAME_PORT = Val(ReadINI("SERVER" & lstServers.ListIndex, "PORT", App.Path & "\Config\Serveur.ini"))
        Me.Caption = "Liste de Serveur - Vérification de l'état du Serveurs..."
        If CheckServerStatus = False Then Me.Caption = "Liste de Serveur - Serveur Hors-Ligne!": cmdOk.Enabled = True: Exit Sub
        cmdOk.Enabled = True
        Me.Caption = "Liste de Serveur - Connecté au Serveur!"
        frmMirage.Socket.Close
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT
        frmMainMenu.Show
        Unload Me
    End Sub

    Private Sub Form_Load()
    Dim i As Long, c As Long
        frmServerChooser.Visible = True

        filename = App.Path & "\Config\Serveur.ini"
        i = 0
        c = 0
        CHECK_WAIT = False
        lstServers.Clear
       
        cmdOk.Enabled = False
        Me.Caption = "Liste de Serveur - Vérification de l'état des Serveurs..."
        Do Until c = 1
            DoEvents
            If CHECK_WAIT = False Then
                If ReadINI("SERVER" & i, "IP", filename) <> vbNullString And ReadINI("SERVER" & i, "PORT", filename) <> vbNullString Then
                    GAME_IP = ReadINI("SERVER" & i, "IP", filename)
                    GAME_PORT = Val(ReadINI("SERVER" & i, "PORT", filename))
                    If CheckServerStatus = True Then CHECK_WAIT = True: Call SendData("serverresults" & SEP_CHAR & i & SEP_CHAR & END_CHAR) Else lstServers.AddItem ReadINI("SERVER" & i, "Name", filename) & " - Fermé!"
                    i = i + 1
                Else
                    c = 1
                End If
            End If
            Sleep 1
        Loop
        cmdOk.Enabled = True
        Me.Caption = "Liste de Serveur - Vérification Terminée!"
        
        Call WriteINI("UPDATER", "exename", App.EXEName, App.Path & "\Config\Updater.ini")
      End Sub

    Function CheckServerStatus() As Boolean
        frmMirage.Socket.Close
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT

        cmdOk.Enabled = False
        CheckServerStatus = False
       
        If ConnectToServer Then
            CheckServerStatus = True
        End If
    End Function

Private Sub Label1_Click()
    Call cmdCancel_Click
End Sub

Private Sub lstServers_DblClick()
    If cmdOk.Enabled = True Then Call cmdOk_Click
End Sub
