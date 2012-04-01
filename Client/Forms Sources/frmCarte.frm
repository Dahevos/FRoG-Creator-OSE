VERSION 5.00
Begin VB.Form frmCarte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carte"
   ClientHeight    =   8505
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8490
   Icon            =   "frmCarte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgcarte 
      Height          =   8505
      Left            =   0
      Picture         =   "frmCarte.frx":17D2A
      Top             =   0
      Width           =   8505
   End
End
Attribute VB_Name = "frmCarte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    For i = 1 To 4
        If i = 1 Then Ending = ".gif"
        If i = 2 Then Ending = ".jpg"
        If i = 3 Then Ending = ".png"
        If i = 4 Then Ending = ".bmp"
        
        If FileExiste("images\Carte" & Ending) Then imgcarte.Picture = LoadPicture(App.Path & "\images\Carte" & Ending)
    Next i
End Sub


