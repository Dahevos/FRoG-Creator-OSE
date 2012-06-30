VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Nouveau Personnage"
   ClientHeight    =   3975
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   3390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   226
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2520
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   20
      Top             =   2400
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   21
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H0080FF80&
      Caption         =   "Femme"
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
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H0080FF80&
      Caption         =   "Homme"
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
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   1080
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5880
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   120
      Top             =   0
   End
   Begin VB.ComboBox cmbClass 
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
      Height          =   300
      ItemData        =   "frmNewChar.frx":2C02A
      Left            =   240
      List            =   "frmNewChar.frx":2C02C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1920
      Width           =   2835
   End
   Begin VB.TextBox txtName 
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
      Height          =   255
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   720
      Width           =   2865
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   0
      Width           =   375
   End
   Begin VB.Label picAddChar 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblSPEED 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   720
      TabIndex        =   19
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Vitesse"
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
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   540
   End
   Begin VB.Label lblDEF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   720
      TabIndex        =   17
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Def"
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
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label lblSTR 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   690
      TabIndex        =   15
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Force"
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
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblMAGI 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Magie"
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
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblSP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "End"
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
      Height          =   240
      Left            =   1440
      TabIndex        =   8
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label lblMP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PM"
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
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1785
      TabIndex        =   5
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PV"
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
      Height          =   195
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sexe"
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
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public animi As Long
Private l As Long

Private Sub cmbClass_Click()
Dim i As Long
l = 0
    For i = 0 To Max_Classes
        If Trim$(cmbClass.List(cmbClass.ListIndex)) = Trim$(Class(i).name) Then l = i: Exit For
    Next i
    
    If l < 0 Or l > Max_Classes Then l = 0
    
    lblHP.Caption = STR$(Class(l).HP)
    lblMP.Caption = STR$(Class(l).MP)
    lblSP.Caption = STR$(Class(l).SP)
    
    lblSTR.Caption = STR$(Class(l).STR)
    lblDEF.Caption = STR$(Class(l).DEF)
    lblSPEED.Caption = STR$(Class(l).speed)
    lblMAGI.Caption = STR$(Class(l).MAGI)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = True
drx = x
dry = y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If dr Then DoEvents: If dr Then Call Me.Move(Me.Left + (x - drx), Me.Top + (y - dry))
If Me.Left > Screen.Width Or Me.Top > Screen.height Then Me.Top = Screen.height \ 2: Me.Left = Screen.Width \ 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
dr = False
drx = 0
dry = 0
End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)
        
        If Len(Trim$(txtName.Text)) < 3 Then MsgBox "Le nom de votre personne doit contenir plus de trois lettres.": Exit Sub
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then Call MsgBox("Vous ne pouvez utiliser de carateres spéciales dans votre nom.", vbOKOnly, GAME_NAME): txtName.Text = vbNullString: Exit Sub
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picAddChar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Msg As String
Dim i As Long

    If Trim$(txtName.Text) <> vbNullString Then
        Msg = Trim$(txtName.Text)
        
        If Len(Trim$(txtName.Text)) < 3 Then MsgBox "Le nom de votre personne doit contenir plus de trois lettres.": Exit Sub
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid$(Msg, i, 1)) < 32 Or Asc(Mid$(Msg, i, 1)) > 126 Then Call MsgBox("Vous ne pouvez utiliser de carateres spéciales dans votre nom.", vbOKOnly, GAME_NAME): txtName.Text = vbNullString: Exit Sub
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmMainMenu.fraPers.Visible = True
    Me.Visible = False
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste(Rep_Theme & "\Login\nouveau_personnage" & Ending) Then Me.Picture = LoadPNG(App.Path & Rep_Theme & "\Login\nouveau_personnage" & Ending)
    Next i
    
    'Picsprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png", True)
    picAddChar.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picAddChar.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picCancel.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picCancel.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    dr = False
    Call ChrgSpriteSurf
End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    frmMainMenu.Visible = True
    frmMainMenu.fraPers.Visible = True
    Me.Visible = False
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 3 Then animi = 0
If l < 0 Or cmbClass.ListIndex < 0 Then Exit Sub
If optMale.Value = True Then
    PrepareSprite (Class(l).MaleSprite)
    Call AffSurfPic(DD_SpriteSurf(Class(l).MaleSprite), Picpic, animi * PIC_X, 0)
Else
    PrepareSprite (Class(l).FemaleSprite)
    Call AffSurfPic(DD_SpriteSurf(Class(l).FemaleSprite), Picpic, animi * PIC_X, 0)
End If
End Sub

Private Sub txtName_GotFocus()
txtName.SelStart = 0
txtName.SelLength = Len(txtName.Text)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call picAddChar_Click
End Sub
