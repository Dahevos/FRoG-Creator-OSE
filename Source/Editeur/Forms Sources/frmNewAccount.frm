VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Nouveau Compte"
   ClientHeight    =   5835
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   5115
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
   Picture         =   "frmNewAccount.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword2 
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
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2640
      Width           =   2685
   End
   Begin VB.TextBox txtPassword 
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
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2130
      Width           =   2685
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
      Top             =   1665
      Width           =   2685
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Retaper le mot  de passe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de compte:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   1770
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   1920
      Width           =   1590
   End
   Begin VB.Label picConnect 
      BackStyle       =   0  'Transparent
      Caption         =   "Créer un compte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label picCancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revenir"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   5400
      Width           =   555
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
Dim Ending As String
    For i = 1 To 3
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
 
        If FileExiste("GUI\NewAccount" & Ending) Then frmNewAccount.Picture = LoadPicture(App.Path & "\GUI\NewAccount" & Ending)
    Next i
    Label1.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label2.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label2.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label3.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label3.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picCancel.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picCancel.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picConnect.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmNewAccount.Visible = False
End Sub

Private Sub picConnect_Click()
Dim Msg As String
Dim i As Long
    
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" And Trim(txtPassword2.Text) <> "" Then
        Msg = Trim(txtName.Text)
        
        If Trim(txtPassword.Text) <> Trim(txtPassword2.Text) Then
            MsgBox "Le mot de passe ne correspond pas!"
            Exit Sub
        End If
        
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Votre nom et mot de passe doit contenir plus de 3 caractères."
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call MsgBox("Vous ne pouvez pas utiliser des accents dans votre noms.", vbOKOnly, GAME_NAME)
                txtName.Text = ""
                Exit Sub
            End If
        Next i
    
        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
End Sub

