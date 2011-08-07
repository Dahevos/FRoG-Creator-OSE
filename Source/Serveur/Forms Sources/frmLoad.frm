VERSION 5.00
Object = "{463051F7-93F6-433B-8C04-1B5EF7493179}#1.0#0"; "WinXPCEngine.ocx"
Begin VB.Form frmLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FRoG Server 0.6"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   0
      Top             =   840
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
      ListBoxControl  =   0   'False
      PictureControl  =   0   'False
      ButtonControl   =   0   'False
      CheckControl    =   0   'False
      OptionControl   =   0   'False
      ComboBoxControl =   0   'False
      DriveListBoxControl=   0   'False
      TabStripControl =   0   'False
      StatusBarControl=   0   'False
      ProgressBarControl=   0   'False
      SliderControl   =   0   'False
      ImageComboControl=   0   'False
      ListViewControl =   0   'False
      FileListBoxControl=   0   'False
      DirListBoxControl=   0   'False
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Initialisation ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Show
    DoEvents
    Dim t As Currency
    If InDestroy = False Then Call InitServer
End Sub

Private Sub Timer1_Timer()
If InDestroy = True Then Unload Me
End Sub
