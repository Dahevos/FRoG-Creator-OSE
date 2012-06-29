VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmenvoier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eléments à envoyer au serveur"
   ClientHeight    =   3360
   ClientLeft      =   270
   ClientTop       =   300
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Envoyer les éléments sélectionnés"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   2655
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2295
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4048
      _Version        =   393217
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Que voulez vous envoyer?"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmenvoier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim i As Long
Dim n As Long
Dim r As Long
Dim z As String

r = 0
For n = 1 To TreeView1.Nodes.Count
    If Not TreeView1.Nodes.Item(n).Checked Then r = r + 1
Next n
If r = Val(TreeView1.Nodes.Count) Then MsgBox "Aucun éléments sélectionés.": Exit Sub

Command1.Visible = False
ProgressBar1.Visible = True
ProgressBar1.value = ProgressBar1.Min
i = TreeView1.Nodes.Count
r = (100 \ i) - 1
For n = 1 To i
    DoEvents
    If n > TreeView1.Nodes.Count Then Exit For
    ProgressBar1.value = ProgressBar1.value + r
    If TreeView1.Nodes(n).Checked = True Then
        z = TreeView1.Nodes(n).Text
        If Mid$(z, 1, Len(z) - 1) = "Objet" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveItem(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "objet" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
        
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 7)) = "Magasin" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveShop(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "magasin" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
        
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 4)) = "Sort" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveSpell(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "sort" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
        
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 3)) = "PNJ" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveNpc(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "pnj" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
        
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 6)) = "Flêche" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveArrow(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "flêche" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
            
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 8)) = "Emoticon" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveEmoticon(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "emot" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
        
        ElseIf Mid$(z, 1, Len(z) - (Len(z) - 5)) = "Quête" And Mid$(z, Len(z)) <> "s" Then
            Call SendSaveQuete(Val(TreeView1.Nodes(n).Tag))
            Call WriteINI("modif", "quête" & Val(TreeView1.Nodes(n).Tag), "0", App.Path & "\config.ini")
            
        End If
    End If
    DoEvents
Next n
Call MsgBox("Envoies terminés.", vbInformation)
If TreeView1.Nodes.Count = 7 Then MsgBox "Plus aucun fichier à envoyer.", vbInformation:  Call TreeView1.Nodes.Clear: Unload Me: Exit Sub
ProgressBar1.value = ProgressBar1.Max
Call TreeView1.Nodes.Clear
Me.Visible = False
Call EnvoieServeur
Command1.Visible = True
ProgressBar1.Visible = False
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim i As Long
If Node.Checked = True Then
    For i = 1 To Node.Children
        TreeView1.Nodes.Item(Node.Index + i).Checked = True
    Next i
End If
End Sub
