VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Nouveau Personnage"
   ClientHeight    =   7440
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewChar.frx":0000
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   572
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3600
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   24
      Top             =   2520
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   25
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FFFFFF&
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
      Left            =   720
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00FFFFFF&
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
      Left            =   720
      TabIndex        =   16
      Top             =   1800
      Value           =   -1  'True
      Width           =   855
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
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   4560
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4560
      Top             =   240
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
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2760
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
      Height          =   225
      Left            =   720
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1485
      Width           =   2760
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
      Left            =   3000
      TabIndex        =   23
      Top             =   3600
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
      Left            =   2520
      TabIndex        =   22
      Top             =   3600
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
      Left            =   1920
      TabIndex        =   21
      Top             =   3600
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
      Left            =   1680
      TabIndex        =   20
      Top             =   3600
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
      Left            =   1080
      TabIndex        =   19
      Top             =   3600
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
      Left            =   720
      TabIndex        =   18
      Top             =   3600
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
      Left            =   3960
      TabIndex        =   15
      Top             =   3360
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
      Left            =   3480
      TabIndex        =   14
      Top             =   3360
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
      Left            =   3000
      TabIndex        =   13
      Top             =   3360
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
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
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
      Left            =   1920
      TabIndex        =   11
      Top             =   3360
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
      Left            =   1680
      TabIndex        =   10
      Top             =   3360
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3360
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
      Left            =   720
      TabIndex        =   8
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vocation :"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du personnage :"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   1230
      Width           =   1530
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Créer un nouveau personnage"
      Height          =   195
      Left            =   1485
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revenir"
      Height          =   195
      Left            =   2370
      TabIndex        =   4
      Top             =   5340
      Width           =   585
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

Private Sub cmbClass_Click()
    lblHP.Caption = STR(Class(cmbClass.ListIndex).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex).SP)
    
    lblSTR.Caption = STR(Class(cmbClass.ListIndex).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex).def)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex).speed)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex).magi)
End Sub

Private Sub picAddChar_Click()
Dim Msg As String
Dim i As Long

    If Trim(txtName.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        If Len(Trim(txtName.Text)) < 3 Then
            MsgBox "Le nom de votre personne doit contenir plus de trois lettres."
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("Vous ne pouvez utiliser de carateres spéciales dans votre nom.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    frmChars.Visible = True
    Me.Visible = False
End Sub



Private Sub Timer1_Timer()
If cmbClass.ListIndex < 0 Then Exit Sub
If optMale.value = True Then
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, Picsprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).MaleSprite) * PIC_Y, SRCCOPY)
Else
    Call BitBlt(picPic.hDC, 0, 0, PIC_X, PIC_Y, Picsprites.hDC, animi * PIC_X, Int(Class(cmbClass.ListIndex).FemaleSprite) * PIC_Y, SRCCOPY)
End If
End Sub

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExiste("GUI\NewCharacter" & Ending) Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\NewCharacter" & Ending)
    Next i
    Picsprites.Picture = LoadPNG(App.Path & "\GFX\sprites.png")
    Label1.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label2.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label2.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label3.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label3.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picAddChar.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picAddChar.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picCancel.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picCancel.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
