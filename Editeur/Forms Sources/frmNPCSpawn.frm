VERSION 5.00
Begin VB.Form frmNPCSpawn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NPC Spawn"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2790
   Icon            =   "frmNPCSpawn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.HScrollBar scrlNum2 
      Height          =   255
      Left            =   240
      Max             =   50
      Min             =   1
      TabIndex        =   4
      Top             =   840
      Value           =   1
      Width           =   2295
   End
   Begin VB.HScrollBar scrlNum3 
      Height          =   255
      Left            =   240
      Max             =   30
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.HScrollBar scrlNum 
      Height          =   255
      Left            =   240
      Min             =   1
      TabIndex        =   0
      Top             =   240
      Value           =   1
      Width           =   2295
   End
   Begin VB.Label lblNum2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de NPC : 1"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label lblNum3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Distance d'apparition (Spawn) : 0"
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
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2385
   End
   Begin VB.Label lblNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numéro du NPC: 1 -"
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1425
   End
End
Attribute VB_Name = "frmNPCSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    NPCSpawnNum = scrlNum.value
    NPCSpawnAmount = scrlNum2.value
    NPCSpawnRange = scrlNum3.value
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    scrlNum2.max = MAX_ATTRIBUTE_NPCS
    
    If NPCSpawnNum < scrlNum.min Then NPCSpawnNum = scrlNum.min
    scrlNum.value = NPCSpawnNum
    If NPCSpawnAmount < scrlNum2.min Then NPCSpawnAmount = scrlNum2.min
    scrlNum2.value = NPCSpawnAmount
    If NPCSpawnRange < scrlNum2.min Then NPCSpawnRange = scrlNum3.min
    scrlNum3.value = NPCSpawnRange
End Sub

Private Sub scrlNum_Change()
    lblNum.Caption = "Numéro du NPC: " & scrlNum.value & " - " & Trim$(Npc(scrlNum.value).name)
End Sub

Private Sub scrlNum2_Change()
    lblNum2.Caption = "Nombre de NPC: " & scrlNum2.value
End Sub

Private Sub scrlNum3_Change()
    lblNum3.Caption = "Distance d'apparition (Spawn): " & scrlNum3.value
End Sub
