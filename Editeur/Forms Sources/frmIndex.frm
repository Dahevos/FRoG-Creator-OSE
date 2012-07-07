VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Éditer..."
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Fermer"
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
      Left            =   3480
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Editer..."
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
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   2415
   End
   Begin VB.ListBox lstIndex 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      ItemData        =   "frmIndex.frx":0000
      Left            =   240
      List            =   "frmIndex.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   3528
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Choix d'Édition"
      TabPicture(0)   =   "frmIndex.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Text1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rechercher (numéro ou nom) :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   2670
         Width           =   2325
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edition"
      Visible         =   0   'False
      Begin VB.Menu couper 
         Caption         =   "Couper"
         Shortcut        =   ^X
      End
      Begin VB.Menu copier 
         Caption         =   "Copier"
         Shortcut        =   ^C
      End
      Begin VB.Menu coller 
         Caption         =   "Coller"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CTRLD As Boolean

Private Sub cmdOk_Click()
    EditorIndex = lstIndex.ListIndex + 1
    
    If EditorIndex <= 0 Then Exit Sub
    
    If InQuetesEditor = True Then
        If HORS_LIGNE = 1 Then Call QuetesEditorInit Else Call SendData("EDITQUETES" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InItemsEditor = True Then
        If HORS_LIGNE = 1 Then Call ItemEditorInit Else Call SendData("EDITITEM" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InNpcEditor = True Then
        If HORS_LIGNE = 1 Then Call NpcEditorInit Else Call SendData("EDITNPC" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InShopEditor = True Then
        If HORS_LIGNE = 1 Then Call ShopEditorInit Else Call SendData("EDITSHOP" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InSpellEditor = True Then
        If HORS_LIGNE = 1 Then Call SpellEditorInit Else Call SendData("EDITSPELL" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InEmoticonEditor = True Then
        If HORS_LIGNE = 1 Then Call EmoticonEditorInit Else Call SendData("EDITEMOTICON" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InArrowEditor = True Then
        If HORS_LIGNE = 1 Then Call ArrowEditorInit Else Call SendData("EDITARROW" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InPetsEditor = True Then
        If HORS_LIGNE = 1 Then Call PetEditorInit Else Call SendData("EDITPET" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InMetierEditor = True Then
        If HORS_LIGNE = 1 Then Call MetierEditorInit Else Call SendData("EDITMETIER" & SEP_CHAR & EditorIndex & END_CHAR)
    ElseIf InRecetteEditor = True Then
        If HORS_LIGNE = 1 Then Call recetteEditorInit Else Call SendData("EDITrecette" & SEP_CHAR & EditorIndex & END_CHAR)
    End If
End Sub

Private Sub cmdCancel_Click()
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InSpellEditor = False
    InEmoticonEditor = False
    InArrowEditor = False
    InQuetesEditor = False
    InMetierEditor = False
    InRecetteEditor = False
    InPetsEditor = False
    DonID = 0
    Unload frmIndex
    frmMirage.SetFocus
End Sub

Private Sub coller_Click()
Dim FileName As String
Dim f As Long
If DonID = lstIndex.ListIndex + 1 Then Exit Sub
    If InQuetesEditor Then
        If FileExist("quetes\quete" & DonID & ".fcq") Then Call FileCopy(App.Path & "\quetes\quete" & DonID & ".fcq", App.Path & "\quetes\quete" & lstIndex.ListIndex + 1 & ".fcq") Else Call SendSaveQuete(DonID): Call FileCopy(App.Path & "\quetes\quete" & DonID & ".fcq", App.Path & "\quetes\quete" & lstIndex.ListIndex + 1 & ".fcq")
        Call ClearQuete(lstIndex.ListIndex + 1)
        If FileExist("quetes\quete" & lstIndex.ListIndex + 1 & ".fcq") Then
            FileName = App.Path & "\quetes\quete" & lstIndex.ListIndex + 1 & ".fcq"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , quete(lstIndex.ListIndex + 1)
            Close #f
        End If
        Call SendSaveQuete(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & quete(lstIndex.ListIndex + 1).nom
    ElseIf InItemsEditor Then
        If FileExist("items\item" & DonID & ".fco") Then Call FileCopy(App.Path & "\items\item" & DonID & ".fco", App.Path & "\items\item" & lstIndex.ListIndex + 1 & ".fco") Else Call SendSaveItem(DonID): Call FileCopy(App.Path & "\items\item" & DonID & ".fco", App.Path & "\items\item" & lstIndex.ListIndex + 1 & ".fco")
        Call ClearItem(lstIndex.ListIndex + 1)
        If FileExist("items\item" & lstIndex.ListIndex + 1 & ".fco") Then
            FileName = App.Path & "\items\item" & lstIndex.ListIndex + 1 & ".fco"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Item(lstIndex.ListIndex + 1)
            Close #f
        End If
        Call SendSaveItem(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & Item(lstIndex.ListIndex + 1).name
    ElseIf InNpcEditor Then
        If FileExist("pnjs\npc" & DonID & ".fcp") Then Call FileCopy(App.Path & "\pnjs\npc" & DonID & ".fcp", App.Path & "\pnjs\npc" & lstIndex.ListIndex + 1 & ".fcp") Else Call SendSaveNpc(DonID): Call FileCopy(App.Path & "\pnjs\npc" & DonID & ".fcp", App.Path & "\pnjs\npc" & lstIndex.ListIndex + 1 & ".fcp")
        Call ClearNpc(lstIndex.ListIndex + 1)
        If FileExist("pnjs\npc" & lstIndex.ListIndex + 1 & ".fcp") Then
            FileName = App.Path & "\pnjs\npc" & lstIndex.ListIndex + 1 & ".fcp"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Npc(lstIndex.ListIndex + 1)
            Close #f
        End If
        Call SendSaveNpc(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & Npc(lstIndex.ListIndex + 1).name
    ElseIf InShopEditor Then
        If FileExist("shops\shop" & DonID & ".fcm") Then Call FileCopy(App.Path & "\shops\shop" & DonID & ".fcm", App.Path & "\shops\shop" & lstIndex.ListIndex + 1 & ".fcm") Else Call SendSaveShop(DonID): Call FileCopy(App.Path & "\shops\shop" & DonID & ".fcm", App.Path & "\shops\shop" & lstIndex.ListIndex + 1 & ".fcm")
        Call ClearShop(lstIndex.ListIndex + 1)
        If FileExist("shops\shop" & lstIndex.ListIndex + 1 & ".fcm") Then
            FileName = App.Path & "\shops\shop" & lstIndex.ListIndex + 1 & ".fcm"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Shop(lstIndex.ListIndex + 1)
            Close #f
        End If
        Call SendSaveShop(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & Shop(lstIndex.ListIndex + 1).name
    ElseIf InSpellEditor Then
        If FileExist("spells\spells" & DonID & ".fcg") Then Call FileCopy(App.Path & "\spells\spells" & DonID & ".fcg", App.Path & "\spells\spells" & lstIndex.ListIndex + 1 & ".fcg") Else Call SendSaveSpell(DonID): Call FileCopy(App.Path & "\spells\spells" & DonID & ".fcg", App.Path & "\spells\spells" & lstIndex.ListIndex + 1 & ".fcg")
        Call ClearSpell(lstIndex.ListIndex + 1)
        If FileExist("spells\spells" & lstIndex.ListIndex + 1 & ".fcg") Then
            FileName = App.Path & "\spells\spells" & lstIndex.ListIndex + 1 & ".fcg"
            f = FreeFile
            Open FileName For Binary As #f
                Get #f, , Spell(lstIndex.ListIndex + 1)
            Close #f
        End If
        Call SendSaveSpell(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & Spell(lstIndex.ListIndex + 1).name
    ElseIf InEmoticonEditor Then
        Emoticons(lstIndex.ListIndex).Command = Trim$(Emoticons(DonID - 1).Command)
        Emoticons(lstIndex.ListIndex).Pic = Val(Emoticons(DonID - 1).Pic)
        Call SendSaveEmoticon(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex & " : " & Emoticons(lstIndex.ListIndex).Command
    ElseIf InArrowEditor Then
        Arrows(lstIndex.ListIndex + 1).name = Trim$(Arrows(DonID).name)
        Arrows(lstIndex.ListIndex + 1).Pic = Val(Arrows(DonID).Pic)
        Arrows(lstIndex.ListIndex + 1).Range = Arrows(DonID).Range
        Call SendSaveArrow(lstIndex.ListIndex + 1)
        lstIndex.List(lstIndex.ListIndex) = lstIndex.ListIndex + 1 & " : " & Arrows(lstIndex.ListIndex + 1).name
    End If
    
    If DonTP = 1 Then
        If InQuetesEditor Then
            If FileExist("quetes\quete" & DonID & ".fcq") Then Call Kill(App.Path & "\quetes\quete" & DonID & ".fcq")
            Call ClearQuete(DonID)
            Call SendSaveQuete(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InItemsEditor Then
            If FileExist("items\item" & DonID & ".fco") Then Call Kill(App.Path & "\items\item" & DonID & ".fco")
            Call ClearItem(DonID)
            Call SendSaveItem(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InNpcEditor Then
            If FileExist("pnjs\npc" & DonID & ".fcp") Then Call Kill(App.Path & "\pnjs\npc" & DonID & ".fcp")
            Call ClearNpc(DonID)
            Call SendSaveNpc(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InShopEditor Then
            If FileExist("shops\shop" & DonID & ".fcm") Then Call Kill(App.Path & "\shops\shop" & DonID & ".fcm")
            Call ClearShop(DonID)
            Call SendSaveShop(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InSpellEditor Then
            If FileExist("spells\spells" & DonID & ".fcg") Then Call Kill(App.Path & "\spells\spells" & DonID & ".fcg")
            Call ClearSpell(DonID)
            Call SendSaveSpell(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InEmoticonEditor Then
            Emoticons(DonID - 1).Command = vbNullString
            Emoticons(DonID - 1).Pic = 0
            Call SendSaveEmoticon(DonID)
            lstIndex.List(DonID - 1) = DonID - 1 & " : "
        ElseIf InArrowEditor Then
            Arrows(DonID).name = vbNullString
            Arrows(DonID).Pic = 0
            Arrows(DonID).Range = 0
            Call SendSaveArrow(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        ElseIf InPetsEditor Then
            If FileExist("pets\pet" & DonID & ".fcf") Then Call Kill(App.Path & "\pets\pet" & DonID & ".fcf")
            Call ClearPet(DonID)
            Call SendSavePet(DonID)
            lstIndex.List(DonID - 1) = DonID & " : "
        End If
    End If
End Sub

Private Sub copier_Click()
    DonID = lstIndex.ListIndex + 1
    DonTP = 2
End Sub

Private Sub couper_Click()
    DonID = lstIndex.ListIndex + 1
    DonTP = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
If KeyCode = vbKeyControl Then CTRLD = True
If CTRLD And KeyCode = vbKeyC Then Call copier_Click
If CTRLD And KeyCode = vbKeyV Then Call coller_Click
If CTRLD And KeyCode = vbKeyX Then Call couper_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then CTRLD = False
End Sub

Private Sub lstIndex_DblClick()
Call cmdOk_Click
End Sub

Private Sub lstIndex_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
End Sub

Private Sub lstIndex_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    If DonID > 0 Then coller.Enabled = True Else coller.Enabled = False
    Call PopupMenu(edit)
End If
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call cmdOk_Click
End Sub

Private Sub Text1_Change()
On Error GoTo er:
If Trim$(Text1.Text) = vbNullString Then lstIndex.ListIndex = 0: Exit Sub
If IsNumeric(Text1.Text) Then
    lstIndex.ListIndex = Val(Text1.Text) - 1
Else
    Dim i As Long
    For i = 0 To lstIndex.ListCount
        If InStr(1, lstIndex.List(i), Trim$(Text1.Text)) Then lstIndex.ListIndex = i
    Next i
End If
Exit Sub
er:
MsgBox "Numéro ou Nom introuvable.", vbCritical
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub
