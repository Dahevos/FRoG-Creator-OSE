VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainMenu 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Principal"
   ClientHeight    =   7200
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   9585
   ClipControls    =   0   'False
   ForeColor       =   &H000000FF&
   Icon            =   "frmMainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Palette         =   "frmMainMenu.frx":17D2A
   Picture         =   "frmMainMenu.frx":1C7EB
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPicChar 
      Interval        =   500
      Left            =   5760
      Top             =   0
   End
   Begin VB.Frame fraPers 
      Caption         =   "Frame1"
      Height          =   3600
      Left            =   5520
      TabIndex        =   20
      Top             =   600
      Width           =   3390
      Begin VB.PictureBox PicChar 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   360
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   29
         Top             =   1680
         Width           =   960
      End
      Begin VB.ListBox lstChars 
         Appearance      =   0  'Flat
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
         Height          =   930
         ItemData        =   "frmMainMenu.frx":3706B
         Left            =   120
         List            =   "frmMainMenu.frx":3706D
         TabIndex        =   6
         Top             =   640
         Width           =   3135
      End
      Begin VB.Label lblCharClasse 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label lblCharLvl 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   1560
         TabIndex        =   27
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblCharNom 
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label picUseChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   435
         Left            =   1440
         TabIndex        =   8
         Top             =   3240
         Width           =   1845
      End
      Begin VB.Label picDelChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   435
         Left            =   240
         TabIndex        =   9
         Top             =   3240
         Width           =   1245
      End
      Begin VB.Label picNewChar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2880
         Width           =   3285
      End
      Begin VB.Label picCancel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   3120
         TabIndex        =   21
         Top             =   0
         Width           =   300
      End
      Begin VB.Image imgPers 
         Height          =   3600
         Left            =   0
         MousePointer    =   5  'Size
         Picture         =   "frmMainMenu.frx":3706F
         Top             =   0
         Width           =   3390
      End
   End
   Begin MSComctlLib.ImageList imgl 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   65280
      _Version        =   393216
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Frame1"
      Height          =   2265
      Left            =   5520
      TabIndex        =   16
      Top             =   600
      Width           =   3390
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Musique"
         Height          =   195
         Left            =   2160
         TabIndex        =   25
         Top             =   1560
         Width           =   180
      End
      Begin VB.TextBox txtName 
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
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         Top             =   600
         Width           =   3075
      End
      Begin VB.TextBox txtPassword 
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   3075
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Save Password"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   195
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Musique"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label picConnect 
         AutoSize        =   -1  'True
         BackColor       =   &H00789298&
         BackStyle       =   0  'Transparent
         Caption         =   "                             "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1920
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memoriser"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   240
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   1005
      End
      Begin VB.Label lbl_creer 
         BackStyle       =   0  'Transparent
         Caption         =   "                                                                              "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Image imgLogin 
         Height          =   2265
         Left            =   0
         MousePointer    =   5  'Size
         Picture         =   "frmMainMenu.frx":5EE31
         Top             =   0
         Width           =   3390
      End
   End
   Begin VB.Timer splash 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7200
      Top             =   0
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmr2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   0
   End
   Begin VB.CheckBox chk_fullscreen 
      BackColor       =   &H80000009&
      Caption         =   "Plein ecran"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Timer Tmrmusic 
      Interval        =   1000
      Left            =   6720
      Top             =   0
   End
   Begin VB.Frame fraNewAccount 
      Caption         =   "Frame2"
      Height          =   3150
      Left            =   5520
      TabIndex        =   18
      Top             =   600
      Width           =   3385
      Begin VB.TextBox txtname2 
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
         Height          =   330
         Left            =   240
         MaxLength       =   20
         TabIndex        =   10
         Top             =   780
         Width           =   2685
      End
      Begin VB.TextBox txtpassword22 
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1560
         Width           =   2685
      End
      Begin VB.TextBox txtPassword2 
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   2280
         Width           =   2685
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "   "
         Height          =   375
         Left            =   3000
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "                           "
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Image imgNouveau 
         Height          =   3150
         Left            =   0
         MousePointer    =   5  'Size
         Picture         =   "frmMainMenu.frx":77F8B
         Top             =   0
         Width           =   3390
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Plein ecran"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label versionlbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   6960
      Width           =   690
   End
   Begin VB.Label picQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quitter"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8400
      TabIndex        =   15
      Top             =   6840
      Width           =   1305
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long
Public DragImg As Long
Public DragX As Long
Public DragY As Long
Private twippx As Long
Private twippy As Long

Public Function getreselotionX()
    getreselotionX = Screen.Width \ Screen.TwipsPerPixelX
End Function

Public Function getreselotionY()
    getreselotionY = Screen.height \ Screen.TwipsPerPixelY
End Function

Private Sub Check1_Click()
If Check1.Value = "0" Then StopMidi Else If FileExiste("Music\mainmenu.mid") Then Call PlayMidi("mainmenu.mid") Else Call PlayMidi("mainmenu.mp3")

Call WriteINI("CONFIG", "Music", STR$(Check1.Value), App.Path & "\Config\Client.ini")
End Sub

Private Sub Form_GotFocus()
If frmNewChar.Visible Then Call frmNewChar.SetFocus
End Sub

'Private Sub chk_fullscreen_Click()
'    If chk_fullscreen.value = "1" Then
'        Call WriteINI("PLEIN_ECRAN", "actif", "1", App.Path & "\Data.ini")
'        frmMirage.Height = Screen.Height / Screen.TwipsPerPixelY
'        frmMirage.Width = Screen.Width / Screen.TwipsPerPixelX
'        frmMirage.picScreen.Height = Screen.Height / Screen.TwipsPerPixelY
'        frmMirage.picScreen.Width = Screen.Width / Screen.TwipsPerPixelX
'        'ChangeScreenSettings 640, 480, 16
'        'Me.WindowState = "2"
'        'frmMainMenu.BorderStyle = "0"
'    Else
'        Call WriteINI("PLEIN_ECRAN", "actif", "0", App.Path & "\Data.ini")
'        frmMirage.Height = 599 * Screen.TwipsPerPixelY
'        frmMirage.Width = 804 * Screen.TwipsPerPixelX
'        frmMirage.picScreen.Height = 599 * Screen.TwipsPerPixelY
'        frmMirage.picScreen.Width = 804 * Screen.TwipsPerPixelX
'        'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
'        frmMirage.WindowState = "0"
'    End If
'End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
On Error Resume Next
    dragAndDrop = 0
    Call iniOptTouche
    charSelectNum = 1
    Check1.Value = Val(ReadINI("CONFIG", "Music", App.Path & "\Config\Client.ini"))
    
    If getreselotionY < 768 Then
    netbook = True
    End If
    
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
 
        If FileExiste(Rep_Theme & "\Login\connexion" & Ending) Then imgLogin.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\connexion" & Ending)
        If FileExiste(Rep_Theme & "\Login\nouveau" & Ending) Then imgNouveau.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\nouveau" & Ending)
        If FileExiste(Rep_Theme & "\Login\personnage" & Ending) Then imgPers.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\personnage" & Ending)
        If FileExiste(Rep_Theme & "\Login\fond" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\fond" & Ending)
        If FileExiste("GFX/Sprites/Sprites0" & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites0" & Ending)
        End If
    Next i
    
    If Check1.Value = 1 Then If FileExiste("Music\mainmenu.mid") Then Call PlayMidi("mainmenu.mid") Else Call PlayMidi("mainmenu.mp3")
            
    'Picsprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png", True)
    
    fraNewAccount.Visible = False
    fraPers.Visible = False
    txtName.Text = Trim$(ReadINI("INFO", "Account", App.Path & "\Config\Account.ini"))
    txtPassword.Text = Trim$(ReadINI("INFO", "Password", App.Path & "\Config\Account.ini"))
    
    If Trim$(txtPassword.Text) <> vbNullString Then Check2.Value = Checked Else Check2.Value = Unchecked
    
    fraLogin.Visible = True
    txtName.SelStart = 0
    txtName.SelLength = Len(txtName)
        
'    If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 0 Then
'        chk_fullscreen.value = "0"
'    Else
'        chk_fullscreen.value = "1"
'    End If
    
    'If ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini") = 1 Then
    'ChangeScreenSettings 640, 480, 16
    'End If
    twippy = Screen.TwipsPerPixelY
    twippx = Screen.TwipsPerPixelX
    
    versionlbl.Caption = "Version: " & ReadINI("VERSION", "Version", (App.Path & "\Config\info.ini"))
    
    Me.Icon = frmMirage.Icon

    fraLogin.Visible = True
    
    Call netbook_change
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call GameDestroy
End Sub

Private Sub imgLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 1
DragX = x
DragY = y
End Sub

Private Sub imgLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 1 Then fraLogin.Top = fraLogin.Top + ((y / twippy) - (DragY / twippy)): fraLogin.Left = fraLogin.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub imgLogin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub imgNouveau_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 2
DragX = x
DragY = y
End Sub

Private Sub imgNouveau_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 2 Then fraNewAccount.Top = fraNewAccount.Top + ((y / twippy) - (DragY / twippy)): fraNewAccount.Left = fraNewAccount.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub imgNouveau_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
DragImg = 0
DragX = 0
DragY = 0
End Sub

Private Sub imgPers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragImg = 3
    DragX = x
    DragY = y
End Sub

Private Sub imgPers_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If DragImg = 3 Then fraPers.Top = fraPers.Top + ((y / twippy) - (DragY / twippy)): fraPers.Left = fraPers.Left + ((x / twippx) - (DragX / twippx))
End Sub

Private Sub imgPers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DragImg = 0
    DragX = 0
    DragY = 0
End Sub

Private Sub Label1_Click()
Dim Msg As String
Dim i As Long
    
    If Trim$(txtname2.Text) <> vbNullString And Trim$(txtpassword22.Text) <> vbNullString And Trim$(txtPassword2.Text) <> vbNullString Then
        Msg = Trim$(txtname2.Text)
        
        If Trim$(txtpassword22.Text) <> Trim$(txtPassword2.Text) Then MsgBox "Le mot de passe ne correspond pas.": Exit Sub
        
        If Len(Trim$(txtname2.Text)) < 3 Or Len(Trim$(txtpassword22.Text)) < 3 Then MsgBox "Votre nom et mot de passe doit contenir plus de 3 caractères.": Exit Sub
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then Call MsgBox("Vous ne pouvez pas utiliser d'accents dans votre nom.", vbOKOnly, GAME_NAME): txtName.Text = vbNullString: Exit Sub
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

Private Sub Label2_Click()
    fraLogin.Visible = True
    fraNewAccount.Visible = False
End Sub

Private Sub Label6_Click()
 Call GameDestroy
End Sub

Private Sub lbl_creer_Click()
    fraNewAccount.Visible = True
    fraLogin.Visible = False
End Sub

Private Sub lstChars_Click()
Dim i As Byte
Dim Ending As String
    charSelectNum = lstChars.ListIndex + 1
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending)
        End If
    Next i
    PicChar.height = PicChar.height / 4
    PicChar.Width = PicChar.Width / 4
    If PicChar.Width > 960 Then
        PicChar.Width = 960
    End If
    If PicChar.height > 960 Then
        PicChar.height = 960
    End If
    If PicChar.Width > 480 Then
        PicChar.Left = 840 - PicChar.Width + 480
    Else
        PicChar.Left = 840
    End If

    If charSelect(charSelectNum).name <> "" Then
        lblCharNom.Caption = charSelect(charSelectNum).name
        lblCharLvl.Caption = "Niv. " & charSelect(charSelectNum).level
        lblCharClasse.Caption = charSelect(charSelectNum).classe
    Else
        lblCharNom.Caption = "Slot Libre"
        lblCharLvl.Caption = ""
        lblCharClasse.Caption = ""
    End If
End Sub

Private Sub lstChars_DblClick()
    Call picUseChar_Click
End Sub

Private Sub lstChars_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call picUseChar_Click: KeyAscii = 0
End Sub

Private Sub picCancel_Click()

 
    Call TcpDestroy(1)
    Sleep (2000)
    fraLogin.Visible = True
    fraPers.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim$(txtName.Text) <> vbNullString And Trim$(txtPassword.Text) <> vbNullString Then
        If Len(Trim$(txtName.Text)) < 3 Or Len(Trim$(txtPassword.Text)) < 3 Then MsgBox "Votre nom et votre mot de passe doivent contenir plus de 3 caractéres": Exit Sub
        Call MenuState(MENU_STATE_LOGIN)
        Call WriteINI("INFO", "Account", txtName.Text, (App.Path & "\Config\Account.ini"))
        If Check2.Value = Checked Then Call WriteINI("INFO", "Password", txtPassword.Text, (App.Path & "\Config\Account.ini")) Else Call WriteINI("INFO", "Password", "", (App.Path & "\Config\Account.ini"))
    End If
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement.": Exit Sub

    Value = MsgBox("Es-tu certains de vouloir éffacer ce personnage?", vbYesNo, GAME_NAME)
    
    If Value = vbYes Then Call MenuState(MENU_STATE_DELCHAR)
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Emplacement libre" Then MsgBox "Il y a déjà un personnage à cette emplacement.": Exit Sub
    Call SendData("PICVALUE" & END_CHAR)
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picQuit_Click()
'Dim Pathy As String
'Pathy = App.Path & "\config.ini"
'ChangeScreenSettings ReadINI("CONFIG", "X", Pathy), ReadINI("CONFIG", "Y", Pathy), 32
    Call GameDestroy
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Emplacement libre" Then MsgBox "Il n'y a pas de personnage à cette emplacement.": Exit Sub
    Call SendData("PICVALUE" & END_CHAR)
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Public Sub ChangeScreenSettings(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long

lIndex = 0

Do
    lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
    If lTemp = 0 Then Exit Do
    lIndex = lIndex + 1
    With tDevMode
        If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight And .dmBitsPerPel = lColors Then lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY): Exit Do
    End With
Loop

End Sub

Private Sub splash_Timer()
frmsplash.Visible = False
splash.Enabled = False
End Sub

Private Sub tmr2_Timer()
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 0 Then
    frmMirage.BorderStyle = 3
    frmMirage.WindowState = 0
    'frmMirage.StartUpPosition = 1
End If
If Val(ReadINI("PLEIN_ECRAN", "actif", App.Path & "\Data.ini")) = 1 Then
    frmMirage.BorderStyle = 0
    frmMirage.WindowState = 2
    'frmMirage.StartUpPosition = 2
End If
End Sub

Private Sub Tmrmusic_Timer()
If frmMirage.Mediaplayer.Controls.currentPosition = 200 Then
    If FileExiste("Music\mainmenu.mid") Then Call PlayMidi("mainmenu.mid") Else Call PlayMidi("mainmenu.mp3")
End If
If Me.Visible = False Then Tmrmusic.Enabled = False Else Tmrmusic.Enabled = True
End Sub

Private Sub tmrPicChar_Timer()
Dim i As Byte
Dim Ending As String
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending) Then
            PicChar.Picture = LoadPNG(App.Path & "/GFX/Sprites/Sprites" & charSelect(charSelectNum).sprt & Ending)
        End If
    Next i
    PicChar.height = PicChar.height / 4
    PicChar.Width = PicChar.Width / 4
    If PicChar.Width > 960 Then
        PicChar.Width = 960
    End If
    If PicChar.height > 960 Then
        PicChar.height = 960
    End If
    If PicChar.Width > 480 Then
        PicChar.Left = 840 - PicChar.Width + 480
    Else
        PicChar.Left = 840
    End If
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub

Private Sub txtname2_GotFocus()
txtname2.SelStart = 0
txtname2.SelLength = Len(txtname2)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picConnect_Click
End Sub

Private Sub txtPassword2_GotFocus()
txtPassword2.SelStart = 0
txtPassword2.SelLength = Len(txtPassword2)
End Sub

Private Sub txtpassword22_GotFocus()
txtpassword22.SelStart = 0
txtpassword22.SelLength = Len(txtpassword22)
End Sub

