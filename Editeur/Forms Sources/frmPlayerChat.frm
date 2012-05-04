VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPlayerChat 
   BorderStyle     =   0  'None
   Caption         =   "Discutions"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerChat.frx":0000
   ScaleHeight     =   6855
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   5415
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9551
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmPlayerChat.frx":148A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSay 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   6360
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   525
      UseMnemonic     =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Width           =   645
   End
End
Attribute VB_Name = "frmPlayerChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim s As String
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = True
drx = x
dry = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (x - drx), Me.Top + (y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.Height Then Me.Top = Screen.Height \ 2: Me.Left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub Label2_Click()
    Call SendData("qchat" & END_CHAR)
End Sub

Private Sub txtChat_GotFocus()
    txtSay.SetFocus
End Sub

Private Sub txtSay_GotFocus()
txtSay.SelStart = 0
txtSay.SelLength = Len(txtSay.Text)
End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    Dim s As String
    If (KeyAscii = vbKeyReturn) Then
        KeyAscii = 0
        If Trim$(txtSay.Text) = vbNullString Then Exit Sub
        s = vbNewLine & GetPlayerName(MyIndex) & "> " & Trim$(txtSay.Text)
        txtChat.SelStart = Len(txtChat.Text)
        txtChat.SelColor = QBColor(Black)
        txtChat.SelText = s
        txtChat.SelStart = Len(txtChat.Text) - 1
        
        Call SendData("sendchat" & SEP_CHAR & txtSay.Text & END_CHAR)
        txtSay.Text = vbNullString
    End If
End Sub
