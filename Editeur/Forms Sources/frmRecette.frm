VERSION 5.00
Begin VB.Form frmRecette 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editeur de Recette"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Objet Donné"
      Height          =   855
      Left            =   120
      TabIndex        =   56
      Top             =   5880
      Width           =   6135
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   10
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   61
         Top             =   480
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   10
         Left            =   4800
         TabIndex        =   62
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   10
         Left            =   840
         TabIndex        =   60
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Quantité "
         Height          =   255
         Left            =   3120
         TabIndex        =   58
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Objet"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Objets dans la Recette"
      Height          =   5175
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6135
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   9
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   54
         Top             =   4800
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   52
         Top             =   4800
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   8
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   49
         Top             =   4320
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   47
         Top             =   4320
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   7
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   44
         Top             =   3840
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   3840
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   6
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   39
         Top             =   3360
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   3360
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   5
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   34
         Top             =   2880
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   2880
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   4
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   29
         Top             =   2400
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   2400
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   3
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   24
         Top             =   1920
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   2
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   19
         Top             =   1440
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   1
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   14
         Top             =   960
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemQ 
         Height          =   255
         Index           =   0
         Left            =   3120
         Max             =   200
         Min             =   1
         TabIndex        =   9
         Top             =   480
         Value           =   1
         Width           =   2895
      End
      Begin VB.HScrollBar scrlItemNum1 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   9
         Left            =   4800
         TabIndex        =   55
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   53
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label20 
         Caption         =   "Objet 10"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   8
         Left            =   4800
         TabIndex        =   50
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   48
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Objet 9"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   7
         Left            =   4800
         TabIndex        =   45
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   43
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label16 
         Caption         =   "Objet 8"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   40
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   38
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label14 
         Caption         =   "Objet 7"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   5
         Left            =   4800
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   33
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label12 
         Caption         =   "Objet 6"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   30
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   28
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Objet 5"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   25
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Objet 4"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   20
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Objet 3"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Objet 2"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblItemQ 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblItemNum1 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Quantité "
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Objet 1"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Annuler"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox TxtNom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   6135
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nom de la recette:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "frmRecette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Call recetteEditorCancel
End Sub

Private Sub cmdOk_Click()
    Call recetteEditorOk
End Sub

Private Sub Form_Load()
    Dim i As Byte
    For i = 0 To 10
        scrlItemNum1(i).Max = MAX_ITEMS
    Next i
End Sub

Private Sub scrlItemNum1_Change(Index As Integer)
    If scrlItemNum1(Index).value > 0 Then
        lblItemNum1(Index).Caption = scrlItemNum1(Index).value & ": " & Item(scrlItemNum1(Index).value).name
        scrlItemQ(Index).value = 1
        lblItemQ(Index).Caption = 1
        If Item(scrlItemNum1(Index).value).Type = ITEM_TYPE_CURRENCY Or Item(scrlItemNum1(Index).value).Empilable <> 0 Then
            scrlItemQ(Index).Enabled = True
        Else
            scrlItemQ(Index).Enabled = False
        End If
    Else
        lblItemNum1(Index).Caption = "Pas d'objet"
    End If

End Sub

Private Sub scrlItemQ_Change(Index As Integer)
    lblItemQ(Index).Caption = scrlItemQ(Index).value
End Sub
