Attribute VB_Name = "modDatabase"
Option Explicit

Public SOffsetX As Integer
Public SOffsetY As Integer

Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then StripTerminator = Left$(strString, intZeroPos - 1) Else StripTerminator = strString
End Function

Public Function FileExiste(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    FileExiste = True
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & FileName)) = 0 Then FileExiste = False
    Else
        If LenB(Dir$(FileName)) = 0 Then FileExiste = False
    End If
End Function
Sub SaveLocalMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
                            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".fcc"
        
    If Not FileExiste("maps\map" & MapNum & ".fcc") Then Exit Sub
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Map(MapNum)
    Close #f
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim GlobalX As Integer
Dim GlobalY As Integer

GlobalX = PB.Left
GlobalY = PB.Top

If Button = 1 Then PB.Left = GlobalX + x - SOffsetX: PB.Top = GlobalY + y - SOffsetY
End Sub

Sub MoveForm(f As Form, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim GlobalX As Integer
Dim GlobalY As Integer

GlobalX = f.Left
GlobalY = f.Top

If Button = 1 Then f.Left = GlobalX + x - SOffsetX: f.Top = GlobalY + y - SOffsetY
End Sub
