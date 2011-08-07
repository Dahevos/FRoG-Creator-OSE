VERSION 5.00
Begin VB.Form frmInfoBulle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmInfoBulle.frx":0000
   ScaleHeight     =   1635
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   3360
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   3840
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   2880
      Top             =   840
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   3600
      Picture         =   "frmInfoBulle.frx":028A
      ScaleHeight     =   570
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin VB.Label msgs 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "text"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serveur de Frog Creator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2040
   End
End
Attribute VB_Name = "frmInfoBulle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call TransRegion(Me, 0, RGB(255, 255, 255))
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE))
    Me.top = Screen.Height - Me.Height - 300
    Me.Left = Screen.Width - Me.Width - 800
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
    ElseIf Button = 1 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        frmServer.WindowState = 0
        frmServer.Show
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
    ElseIf Button = 1 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        frmServer.WindowState = 0
        frmServer.Show
    End If
End Sub

Private Sub msgs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
    ElseIf Button = 1 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        frmServer.WindowState = 0
        frmServer.Show
    End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
    ElseIf Button = 1 Then
        Timer2.Enabled = False
        Timer3.Enabled = True
        frmServer.WindowState = 0
        frmServer.Show
    End If
End Sub

Private Sub Timer1_Timer()
    If Timer3.Enabled = True Then Timer3.Enabled = False
    If Tra = 250 Then Tra = 255: Timer1.Enabled = False: IBVisible = True
    Call TransRegion(Me, Tra, RGB(255, 255, 255))
    If Tra < 255 Then Tra = Tra + 10
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
    If Timer1.Enabled = True Then Timer1.Enabled = False
    If Tra = 5 Then Tra = 0: Timer3.Enabled = False: IBVisible = False: IBCharge = False: frmInfoBulle.Visible = False
    Call TransRegion(Me, Tra, RGB(255, 255, 255))
    If Tra > 0 Then Tra = Tra - 10
End Sub
