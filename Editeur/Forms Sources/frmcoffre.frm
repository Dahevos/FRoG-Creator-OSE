VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmcoffre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Coffre"
   ClientHeight    =   4905
   ClientLeft      =   645
   ClientTop       =   -285
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   370
      TabMaxWidth     =   2646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Configuration"
      TabPicture(0)   =   "frmcoffre.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblItem"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "code"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdOk"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCancel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Check2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "scrlItem"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkTake"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "HScroll1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   960
         Max             =   500
         Min             =   1
         TabIndex        =   14
         Top             =   3720
         Value           =   1
         Width           =   3255
      End
      Begin VB.CheckBox chkTake 
         Caption         =   "Faire disparaitre la clée une fois utilisé."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2880
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.HScrollBar scrlItem 
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         Max             =   500
         Min             =   1
         TabIndex        =   7
         Top             =   2400
         Value           =   1
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Coffre à clée"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Coffre à code"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Annuler"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox code 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Objet :"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Objet donné :"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label lblItem 
         Caption         =   "1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Objet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Objet"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Entrez le code désiré SVP :"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmcoffre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
If Check2.value = Checked Then
    Check2.value = Unchecked
    Label2.Enabled = False
    Label3.Enabled = False
    scrlItem.Enabled = False
    lblName.Enabled = False
    chkTake.Enabled = False
    lblItem.Enabled = False
End If
If Label1.Enabled Then
    Label1.Enabled = False
    code.Enabled = False
Else
    Label1.Enabled = True
    code.Enabled = True
    Check1.value = Checked
End If
On Error Resume Next
Call code.SetFocus
End Sub

Private Sub Check2_Click()
If Check1.value = Checked Then
    Check1.value = Unchecked
    Label1.Enabled = False
    code.Enabled = False
End If
If Label2.Enabled Then
    Label2.Enabled = False
    Label3.Enabled = False
    scrlItem.Enabled = False
    lblName.Enabled = False
    chkTake.Enabled = False
    lblItem.Enabled = False
Else
    Label2.Enabled = True
    Label3.Enabled = True
    scrlItem.Enabled = True
    lblName.Enabled = True
    chkTake.Enabled = True
    lblItem.Enabled = True
    Check2.value = Checked
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub cmdOk_Click()
    ObjCoffreNum = HScroll1.value
    If Check2.value = Checked Then
        CleCoffreNum = scrlItem.value
        CleCoffreSupr = chkTake.value
        CodeCoffre = vbNullString
        Unload Me
    Else
        CleCoffreNum = 0
        CleCoffreSupr = 0
        CodeCoffre = code.Text
        Unload Me
    End If
    On Error Resume Next
frmMirage.txtMyTextBox.SetFocus
End Sub

Private Sub Form_Load()
    scrlItem.Max = MAX_ITEMS
    HScroll1.Max = MAX_ITEMS
    code.Text = CodeCoffre
    If CleCoffreNum < scrlItem.min Then CleCoffreNum = scrlItem.min
    scrlItem.value = CleCoffreNum
    chkTake.value = CleCoffreSupr
    If ObjCoffreNum < HScroll1.min Then HScroll1.value = HScroll1.min Else HScroll1.value = ObjCoffreNum
    lblName.Caption = Trim$(Item(scrlItem.value).name)
    Label6.Caption = CStr(HScroll1.value)
    Label4.Caption = "Objet donner : " & Trim$(Item(HScroll1.value).name)
End Sub

Private Sub HScroll1_Change()
    Label6.Caption = CStr(HScroll1.value)
    Label4.Caption = "Objet donner : " & Trim$(Item(HScroll1.value).name)
End Sub

Private Sub scrlItem_Change()
    lblItem.Caption = CStr(scrlItem.value)
    lblName.Caption = Trim$(Item(scrlItem.value).name)
End Sub
