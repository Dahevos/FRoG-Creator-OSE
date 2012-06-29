VERSION 5.00
Begin VB.Form frmServerChooser 
   BorderStyle     =   0  'None
   Caption         =   "Sélection du Serveur"
   ClientHeight    =   2610
   ClientLeft      =   2805
   ClientTop       =   1365
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
   Begin VB.CommandButton CmdRafraichir 
      Caption         =   "Rafraichir"
      Height          =   255
      Left            =   1860
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
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
      Left            =   3360
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Connecter"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   0
      Width           =   375
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
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10

    Private Sub cmdCancel_Click()
        Call GameDestroy
        Unload Me
    End Sub

    Private Sub cmdOk_Click()
        If lstServers.ListCount <= 0 Then Exit Sub

        GAME_IP = ReadINI("SERVER" & lstServers.ListIndex, "IP", App.Path & "\Config\Serveur.ini")
        GAME_PORT = Val(ReadINI("SERVER" & lstServers.ListIndex, "PORT", App.Path & "\Config\Serveur.ini"))
        Me.Caption = "Liste de Serveur - Vérification de l'état du Serveurs..."
        If Not CheckServerStatus Then Me.Caption = "Liste de Serveur - Serveur Hors-Ligne.": cmdOk.Enabled = True: Exit Sub
        cmdOk.Enabled = True
        Me.Caption = "Liste de Serveur - Connecté au Serveur."
        frmMirage.Socket.Close
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT
        frmMainMenu.Show
        Call frmMainMenu.txtName.SetFocus
        Unload Me
    End Sub

Private Sub CmdRafraichir_Click()
    Call Form_Load
End Sub

    Private Sub Form_Load()
    Dim FileName As String
    Dim i As Byte, C As Byte, n As Byte
    Dim Ending As String
    Dim Auto_IP As String
    Dim Auto_Port As String
    
    frmServerChooser.Visible = False
    
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"

        If FileExiste(Rep_Theme & "\Login\choix_serveur" & Ending) Then frmServerChooser.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\choix_serveur" & Ending)
    Next i

        FileName = App.Path & "\Config\Serveur.ini"
        i = 0
        C = 0
        CHECK_WAIT = False
        lstServers.Clear
       
        cmdOk.Enabled = False
        CmdRafraichir.Enabled = False
        
        Me.MousePointer = 13
        Do Until C = 1
            DoEvents
            If Not CHECK_WAIT Then
                If ReadINI("SERVER" & i, "IP", FileName) <> vbNullString And ReadINI("SERVER" & i, "PORT", FileName) <> vbNullString Then
                    GAME_IP = ReadINI("SERVER" & i, "IP", FileName)
                    GAME_PORT = Val(ReadINI("SERVER" & i, "PORT", FileName))
                    If CheckServerStatus Then
                    Auto_IP = GAME_IP
                    Auto_Port = GAME_PORT
                    CHECK_WAIT = True
                    Call SendData("serverresults" & SEP_CHAR & i & END_CHAR)
                    n = n + 1
                    Else
                    lstServers.AddItem ReadINI("SERVER" & i, "Name", FileName) & " - Fermé."
                    End If
                    i = i + 1
                Else
                    C = 1
                End If
            End If
            Sleep 1
        Loop
        
        If n = 1 Then
        GAME_IP = Auto_IP
        GAME_PORT = Auto_Port
        frmMirage.Socket.Close
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT
        frmMainMenu.Show
        Call frmMainMenu.txtName.SetFocus
        PostMessage Me.hwnd, WM_CLOSE, 0, 0
        Exit Sub
        Else
        frmServerChooser.Visible = True
        End If
   
        
        cmdOk.Enabled = True
        CmdRafraichir.Enabled = True
        Me.MousePointer = 0
        
        Call WriteINI("UPDATER", "exename", App.EXEName, App.Path & "\Config\Updater.ini")
      End Sub

    Function CheckServerStatus() As Boolean
        frmMirage.Socket.Close
        frmMirage.Socket.RemoteHost = GAME_IP
        frmMirage.Socket.RemotePort = GAME_PORT

        cmdOk.Enabled = False
        CheckServerStatus = False
       
        If ConnectToServer = True Then
            CheckServerStatus = True
        End If
    End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = True
drx = x
dry = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.left + (x - drx), Me.Top + (y - dry))
If Me.left > Screen.Width Or Me.Top > Screen.height Then Me.Top = Screen.height \ 2: Me.left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub Label1_Click()
    Call cmdCancel_Click
End Sub

Private Sub lstServers_DblClick()
    If cmdOk.Enabled = True Then Call cmdOk_Click
End Sub
