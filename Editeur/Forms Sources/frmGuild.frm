VERSION 5.00
Begin VB.Form frmGuild 
   BackColor       =   &H00789298&
   BorderStyle     =   0  'None
   Caption         =   "Création de Guilde"
   ClientHeight    =   5235
   ClientLeft      =   30
   ClientTop       =   -60
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGuild.frx":0000
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtGuild 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Command2 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   4560
      TabIndex        =   3
      Top             =   0
      Width           =   555
   End
   Begin VB.Label Command1 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   2160
      TabIndex        =   2
      Top             =   4800
      Width           =   645
   End
End
Attribute VB_Name = "frmGuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Packet As String

Packet = "MAKEGUILD" & SEP_CHAR & txtName.Text & SEP_CHAR & txtGuild.Text & END_CHAR

Call SendData(Packet)
End Sub

Private Sub Command2_Click()
Unload Me
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
