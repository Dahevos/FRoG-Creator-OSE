VERSION 5.00
Begin VB.Form frmDeleteAccount 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Éffacer un compte"
   ClientHeight    =   5835
   ClientLeft      =   150
   ClientTop       =   -30
   ClientWidth     =   5115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDeleteAccount.frx":0000
   ScaleHeight     =   5835
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
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
      Top             =   2280
      Width           =   2460
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
      Top             =   1560
      Width           =   2460
   End
   Begin VB.Label picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Revenir"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      TabIndex        =   5
      Top             =   5400
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du compte:"
      Height          =   195
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mot de passe:"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label picConnect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "   Éffacer le compte"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   2640
      Width           =   1380
   End
End
Attribute VB_Name = "frmDeleteAccount"
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
 
        If FileExiste("GUI\DeleteAccount" & Ending) Then frmDeleteAccount.Picture = LoadPicture(App.Path & "\GUI\DeleteAccount" & Ending)
    Next i
    Label1.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label2.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label2.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picCancel.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picCancel.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picConnect.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picConnect.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmDeleteAccount.Visible = False
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> "" And Trim(txtPassword.Text) <> "" Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Votre pseudo ou mot de passe doit faire au minimum trois caracteres"
            Exit Sub
        End If
        Call MenuState(MENU_STATE_DELACCOUNT)
    End If
End Sub

