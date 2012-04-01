VERSION 5.00
Begin VB.Form frmInventory 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mirage Online (Inventory)"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
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
   ScaleHeight     =   4560
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3840
      Picture         =   "frmInventory.frx":0000
      ScaleHeight     =   510
      ScaleWidth      =   3000
      TabIndex        =   5
      Top             =   3960
      Width           =   3000
   End
   Begin VB.PictureBox picDropItem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3840
      Picture         =   "frmInventory.frx":0A07
      ScaleHeight     =   510
      ScaleWidth      =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   3000
   End
   Begin VB.PictureBox picUseItem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   3840
      Picture         =   "frmInventory.frx":14AC
      ScaleHeight     =   510
      ScaleWidth      =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   3000
   End
   Begin VB.ListBox lstInv 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   1920
      ItemData        =   "frmInventory.frx":1EE4
      Left            =   3120
      List            =   "frmInventory.frx":1EE6
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.PictureBox picNewAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3000
      Picture         =   "frmInventory.frx":1EE8
      ScaleHeight     =   825
      ScaleWidth      =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   4800
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "frmInventory.frx":3291
      ScaleHeight     =   4110
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call UpdateInventory
End Sub

Private Sub picUseItem_Click()
    Call SendUseItem(frmInventory.lstInv.ListIndex + 1)
End Sub

Private Sub picDropItem_Click()
Dim Value As Long
Dim InvNum As Long

    InvNum = frmInventory.lstInv.ListIndex + 1

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Show them the drop dialog
            frmDrop.Show vbModal
        Else
            Call SendDropItem(frmInventory.lstInv.ListIndex + 1, 0)
        End If
    End If
End Sub

Private Sub picCancel_Click()
    Unload Me
End Sub

