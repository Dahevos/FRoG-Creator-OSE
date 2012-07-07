VERSION 5.00
Begin VB.Form frmTile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiles"
   ClientHeight    =   10560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   704
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar Defile 
      Height          =   10560
      LargeChange     =   10
      Left            =   6720
      Max             =   300
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   10560
      Left            =   0
      ScaleHeight     =   704
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   448
      TabIndex        =   0
      Top             =   0
      Width           =   6720
      Begin VB.Shape shpSelected 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   0
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim KeyShift As Boolean

Private Sub Defile_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyShift Then KeyShift = True
End Sub

Private Sub Defile_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub Form_GotFocus()
Call AffSurfPic(DD_TileSurf(EditorSet), picTile, 0, Defile.value * PIC_Y)
End Sub

Private Sub Defile_Change()
On Error Resume Next
If (EditorTileY * PIC_Y) < frmTile.picTile.Height + (frmTile.Defile.value * PIC_Y) And (EditorTileY * PIC_Y) > ((frmTile.Defile.value - 1) * PIC_Y) Then frmTile.shpSelected.Top = Int((EditorTileY - frmTile.Defile.value) * PIC_Y): frmTile.shpSelected.Visible = True Else frmTile.shpSelected.Visible = False
shpSelected.Left = Int(EditorTileX * PIC_Y)
Call AffSurfPic(DD_TileSurf(EditorSet), picTile, 0, Defile.value * PIC_Y)
End Sub

Private Sub Defile_Scroll()
On Error Resume Next
If (EditorTileY * PIC_Y) < frmTile.picTile.Height + (frmTile.Defile.value * PIC_Y) And (EditorTileY * PIC_Y) > ((frmTile.Defile.value - 1) * PIC_Y) Then frmTile.shpSelected.Top = Int((EditorTileY - frmTile.Defile.value) * PIC_Y): frmTile.shpSelected.Visible = True Else frmTile.shpSelected.Visible = False
shpSelected.Left = Int(EditorTileX * PIC_Y)
Call AffSurfPic(DD_TileSurf(EditorSet), picTile, 0, Defile.value * PIC_Y)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyShift Then KeyShift = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub Form_Load()
Call Defile_Scroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMirage.scrlPicture.value = Defile.value
End Sub

Private Sub picTile_GotFocus()
Defile.SetFocus
End Sub

Private Sub picTile_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyShift Then KeyShift = True
End Sub

Private Sub picTile_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyShift = False
End Sub

Private Sub picTile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call EditorChooseTiles(Button, Shift, x, y)
            shpSelected.Width = 32
            shpSelected.Height = 32
            frmMirage.shpSelected.Width = 32
            frmMirage.shpSelected.Height = 32
            'If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
        Else
            EditorTileX = (x \ PIC_X)
            EditorTileY = (y \ PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                frmMirage.shpSelected.Width = frmMirage.shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                        frmMirage.shpSelected.Width = frmMirage.shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                frmMirage.shpSelected.Height = frmMirage.shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                        frmMirage.shpSelected.Height = frmMirage.shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
            EditorTileX = (shpSelected.Left \ PIC_X)
            EditorTileY = (shpSelected.Top \ PIC_Y) + Defile.value
        End If
    End If
    
    If frmMirage.tp(2).Checked = True Then shpSelected.Width = 32: shpSelected.Height = 32
    'If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
End Sub

Private Sub picTile_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If KeyShift = False Then
            Call EditorChooseTiles(Button, Shift, x, y)
            shpSelected.Width = 32
            shpSelected.Height = 32
            frmMirage.shpSelected.Width = 32
            frmMirage.shpSelected.Height = 32
            'If frmMirage.previsu.Checked And InEditor And frmMirage.tp(1).Checked And frmMirage.MousePointer <> 99 And frmMirage.MousePointer <> 2 Then Call PreVisua
        Else
            EditorTileX = (x \ PIC_X)
            EditorTileY = (y \ PIC_Y)
            
            If Int(EditorTileX * PIC_X) >= shpSelected.Left + shpSelected.Width Then
                EditorTileX = Int(EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                frmMirage.shpSelected.Width = frmMirage.shpSelected.Width + Int(EditorTileX)
            Else
                If shpSelected.Width > PIC_X Then
                    If Int(EditorTileX * PIC_X) >= shpSelected.Left Then
                        EditorTileX = (EditorTileX * PIC_X + PIC_X) - (shpSelected.Left + shpSelected.Width)
                        shpSelected.Width = shpSelected.Width + Int(EditorTileX)
                        frmMirage.shpSelected.Width = frmMirage.shpSelected.Width + Int(EditorTileX)
                    End If
                End If
            End If
            
            If Int(EditorTileY * PIC_Y) >= shpSelected.Top + shpSelected.Height Then
                EditorTileY = Int(EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                frmMirage.shpSelected.Height = frmMirage.shpSelected.Height + Int(EditorTileY)
            Else
                If shpSelected.Height > PIC_Y Then
                    If Int(EditorTileY * PIC_Y) >= shpSelected.Top Then
                        EditorTileY = (EditorTileY * PIC_Y + PIC_Y) - (shpSelected.Top + shpSelected.Height)
                        shpSelected.Height = shpSelected.Height + Int(EditorTileY)
                        frmMirage.shpSelected.Height = frmMirage.shpSelected.Height + Int(EditorTileY)
                    End If
                End If
            End If
            EditorTileX = (shpSelected.Left \ PIC_X)
            EditorTileY = (shpSelected.Top \ PIC_Y) + Defile.value
        End If
    End If
    
    If frmMirage.tp(2).Checked = True Then shpSelected.Width = 32: shpSelected.Height = 32
End Sub

Private Sub picTile_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then picTile.Cls: Unload Me
End Sub
