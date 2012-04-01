VERSION 5.00
Begin VB.Form frmPerso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paramètres"
   ClientHeight    =   1080
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton annuler 
      Caption         =   "Annuler"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   210
      Width           =   1815
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Carte de téléportation :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frmPerso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

