VERSION 5.00
Begin VB.Form frmMariage 
   Caption         =   "Mariage"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Je le veux"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Configurez votre texte de mariage dans mariage.ini"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   840
      TabIndex        =   4
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom de la femme :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du mari :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
End
Attribute VB_Name = "frmMariage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

txtMari = Text1.Text
txtFemme = Text2.Text

Packet = "MARIER" & SEP_CHAR & txtMari & SEP_CHAR & txtFemme & SEP_CHAR & END_CHAR
Call SendData(Packet)
End Sub

Private Sub Form_Load()
Label3.Caption = ReadINI("INTRODUCTION", "texte", App.Path & "\mariage.ini")
End Sub
