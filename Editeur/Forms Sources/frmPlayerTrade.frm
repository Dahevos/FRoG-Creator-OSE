VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BorderStyle     =   0  'None
   Caption         =   "Commerce"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayerTrade.frx":0000
   ScaleHeight     =   457
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Items2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1200
      Left            =   360
      TabIndex        =   2
      Top             =   4680
      Width           =   7455
   End
   Begin VB.ListBox Items1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   4680
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.ListBox PlayerInv1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1590
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Etat 
      BackStyle       =   0  'Transparent
      Caption         =   "En cours..."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   7680
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4560
      TabIndex        =   5
      Top             =   3360
      Width           =   1035
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   240
      TabIndex        =   4
      Top             =   5880
      Width           =   2355
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   4995
      TabIndex        =   3
      Top             =   5880
      Width           =   2340
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Packet As String
Dim i As Long

    Packet = "swapitems" & SEP_CHAR
    For i = 1 To MAX_PLAYER_TRADES
        Packet = Packet & Trading(i).InvNum & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendData(Packet)

End Sub

Private Sub Command3_Click()
Dim i As Long, n As Long, nb As String
i = PlayerInv1.ListIndex + 1

If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
    For n = 1 To MAX_PLAYER_TRADES
        If Trading(n).InvNum = i Then MsgBox "Vous pouvez échanger cet article seulement une fois.": Exit Sub
        
        If Trading(n).InvNum <= 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                nb = InputBox("Combien de " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " vous-les vous échangez?", "Echange")
                If Val(nb) <= 0 Then
                    MsgBox "Nombre inférieur à 0."
                    Exit Sub
                ElseIf nb > Val(GetPlayerInvItemValue(MyIndex, i)) Then
                    MsgBox "Nombre supérieur au nombre d'objets dans l'inventaire."
                    Exit Sub
                Else
                    PlayerInv1.List(i - 1) = PlayerInv1.Text & "(" & nb & ") **"
                    Items1.List(n - 1) = n & " : " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & "(" & nb & ")"
                    Trading(n).InvNum = i
                    Trading(n).InvName = Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                    Trading(n).InvVal = nb
                    Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & SEP_CHAR & Trading(n).InvVal & END_CHAR)
                    Exit Sub
                End If
            Else
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerPetSlot(MyIndex) = i Then
                    MsgBox "Tu ne peux échanger un objet équipé."
                    Exit Sub
                Else
                    PlayerInv1.List(i - 1) = PlayerInv1.Text & "(1) **"
                    Items1.List(n - 1) = n & " : " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name) & "(1)"
                    Trading(n).InvNum = i
                    Trading(n).InvName = Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).name)
                    Trading(n).InvVal = 1
                    Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & SEP_CHAR & Trading(n).InvVal & END_CHAR)
                    Exit Sub
                End If
            End If
        End If
    Next n
End If
End Sub

Private Sub Command4_Click()
Dim i As Long, n As Long
i = Items1.ListIndex + 1

    If Trading(i).InvNum <= 0 Then
        MsgBox "Pas d'objet a retirer."
        Exit Sub
    End If

    PlayerInv1.List(Trading(i).InvNum - 1) = Mid$(Trim$(PlayerInv1.List(Trading(i).InvNum - 1)), 1, Len(PlayerInv1.List(Trading(i).InvNum - 1)) - Len("(" & CStr(Trading(i).InvVal) & ")**"))
    Items1.List(i - 1) = n & ": <Aucun>"
    Trading(i).InvNum = 0
    Trading(i).InvName = vbNullString
    Call SendData("updatetradeinv" & SEP_CHAR & i & SEP_CHAR & 0 & SEP_CHAR & vbNullString & END_CHAR)
End Sub

Private Sub Command5_Click()
    Call SendData("qtrade" & END_CHAR)
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

Private Sub Items1_DblClick()
Call Command4_Click
End Sub

Private Sub Label1_Click()
    Call SendData("qtrade" & END_CHAR)
End Sub

Private Sub PlayerInv1_DblClick()
Call Command3_Click
End Sub
