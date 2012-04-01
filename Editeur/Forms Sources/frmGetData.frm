VERSION 5.00
Begin VB.Form frmGetData 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Online"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDeleteAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   0
      Picture         =   "frmGetData.frx":0000
      ScaleHeight     =   825
      ScaleWidth      =   4800
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "frmGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
