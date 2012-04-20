VERSION 5.00
Begin VB.Form frmFixItem 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Réparer un objet"
   ClientHeight    =   5835
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   4965
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFixItem.frx":0000
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1935
      Width           =   3465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prix :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4650
      TabIndex        =   3
      Top             =   0
      Width           =   315
   End
   Begin VB.Label chkFix 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1920
      TabIndex        =   2
      Top             =   5280
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sélectionner l'objet à réparer :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   2160
   End
End
Attribute VB_Name = "frmFixItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFix_Click()
    Call SendData("fixitem" & SEP_CHAR & cmbItem.ListIndex + 1 & SEP_CHAR & frmTrade.picFixItems.Tag & END_CHAR)
    Call cmbItem_Click
End Sub

Private Sub cmbItem_Click()
    Dim i As Long
    If GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1) > 0 And GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1) < MAX_ITEMS Then
    If Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data1 > -1 And GetPlayerInvItemDur(MyIndex, cmbItem.ListIndex + 1) > -1 Then
        i = ((Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data1 - GetPlayerInvItemDur(MyIndex, cmbItem.ListIndex + 1)) * (Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data2 / 5)) \ 2
        If i <= 0 Then i = 1
        frmFixItem.Label2.Caption = "Prix : " & i & Item(frmTrade.picFixItems.Tag).name
    Else
        Label2.Caption = "Prix : 0" & Item(frmTrade.picFixItems.Tag).name
    End If
    Else
        Label2.Caption = "Prix : 0" & Item(frmTrade.picFixItems.Tag).name
    End If
End Sub

Private Sub cmbItem_Scroll()
Dim i As Long
    If GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1) > 0 And GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1) < MAX_ITEMS Then
    If Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data1 > -1 And GetPlayerInvItemDur(MyIndex, cmbItem.ListIndex + 1) > -1 Then
        i = ((Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data1 - GetPlayerInvItemDur(MyIndex, cmbItem.ListIndex + 1)) * (Item(GetPlayerInvItemNum(MyIndex, cmbItem.ListIndex + 1)).Data2 / 5)) \ 2
        If i <= 0 Then i = 1
        frmFixItem.Label2.Caption = "Prix : " & i & Item(frmTrade.picFixItems.Tag).name
    Else
        Label2.Caption = "Prix : 0" & Item(frmTrade.picFixItems.Tag).name
    End If
    Else
        Label2.Caption = "Prix : 0" & Item(frmTrade.picFixItems.Tag).name
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

Private Sub picCancel_Click()
    Unload Me
    If frmTrade.Visible Then frmTrade.SetFocus
End Sub

