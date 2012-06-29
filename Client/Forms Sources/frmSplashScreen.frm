VERSION 5.00
Begin VB.Form frmSplashScreen 
   Caption         =   "Crée avec"
   ClientHeight    =   7050
   ClientLeft      =   4995
   ClientTop       =   4095
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer splashtimer 
      Interval        =   1500
      Left            =   6360
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -960
      Picture         =   "frmSplashScreen.frx":0000
      Top             =   -1080
      Width           =   12000
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub splashtimer_Timer()
    frmSplashScreen.Visible = False
    splashtimer.Enabled = False
    Call Main
End Sub
