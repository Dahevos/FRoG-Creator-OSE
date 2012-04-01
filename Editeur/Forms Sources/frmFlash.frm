VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10e.ocx"
Begin VB.Form frmFlash 
   BorderStyle     =   0  'None
   Caption         =   "Évènement de Flash"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFlash.frx":0000
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Check 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   120
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash 
      Height          =   6000
      Left            =   2295
      TabIndex        =   1
      Top             =   1725
      Width           =   8250
      _cx             =   14552
      _cy             =   10583
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "-1"
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12240
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Annuler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10035
      TabIndex        =   0
      Top             =   7800
      Width           =   585
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check_Timer()
    If Flash.CurrentFrame > 0 Then
        If Flash.CurrentFrame >= Flash.TotalFrames - 1 Then
            Flash.FrameNum = 0
            Flash.Stop
            Check.Enabled = False
            WriteINI "CONFIG", "Music", frmoptions.chkmusic.value, App.Path & "\Config\Account.ini"
            AccOpt.Music = CBool(frmoptions.chkmusic.value)
            Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
            WriteINI "CONFIG", "Sound", frmoptions.chksound.value, App.Path & "\Config\Account.ini"
            AccOpt.Sound = CBool(frmoptions.chksound.value)
            Unload Me
        End If
    End If
End Sub

Private Sub Label1_Click()
    Flash.FrameNum = 0
    Flash.Stop
    Check.Enabled = False
    WriteINI "CONFIG", "Music", frmoptions.chkmusic.value, App.Path & "\Config\Account.ini"
    AccOpt.Music = CBool(frmoptions.chkmusic.value)
    Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
    WriteINI "CONFIG", "Sound", frmoptions.chksound.value, App.Path & "\Config\Account.ini"
    AccOpt.Sound = CBool(frmoptions.chksound.value)
    Unload Me
End Sub

Private Sub Label2_Click()
    Flash.FrameNum = 0
    Flash.Stop
    Check.Enabled = False
    WriteINI "CONFIG", "Music", frmoptions.chkmusic.value, App.Path & "\Config\Account.ini"
    AccOpt.Music = CBool(frmoptions.chkmusic.value)
    Call PlayMidi(Trim$(Map(Player(MyIndex).Map).Music))
    WriteINI "CONFIG", "Sound", frmoptions.chksound.value, App.Path & "\Config\Account.ini"
    AccOpt.Sound = CBool(frmoptions.chksound.value)
    Unload Me
End Sub
