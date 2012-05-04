VERSION 5.00
Begin VB.Form frmpet 
   Caption         =   "Crédits"
   ClientHeight    =   5820
   ClientLeft      =   240
   ClientTop       =   345
   ClientWidth     =   4905
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label creditline1 
      BackStyle       =   0  'Transparent
      Caption         =   "Merci a Hinomi pour sa belle banniére"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCredits.frx":0000
      Height          =   1035
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   3390
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      Caption         =   " Revenir au menu"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   5040
      Width           =   2640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remerciement: Coke, GodSentdeath, Katsuo, Edouard,Dahevos et a toute la communauté de FRoG Creator"
      Height          =   675
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "frmpet"
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

        If FileExiste("GUI\Credits" & Ending) Then frmCredits.Picture = LoadPicture(App.Path & "\GUI\Credits" & Ending)
    Next i
    Label3.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label3.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    Label1.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    Label1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    picCancel.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    picCancel.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
    creditline1.Font = ReadINI("POLICE", "Police", (App.Path & "\Config\Ecriture.ini"))
    creditline1.FontSize = ReadINI("POLICE", "PoliceSize", (App.Path & "\Config\Ecriture.ini"))
End Sub

Private Sub picCancel_Click()
If frmMirage.Visible = True Then
frmCredits.Visible = False
frmMirage.SetFocus
Else
    frmMainMenu.Visible = True
    frmCredits.Visible = False
End If
End Sub

