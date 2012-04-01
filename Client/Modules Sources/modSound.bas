Attribute VB_Name = "modSound"
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_SYNC As Long = &H0
Public Const SND_ASYNC As Long = &H1
Public Const SND_NODEFAULT As Long = &H2
Public Const SND_MEMORY As Long = &H4
Public Const SND_LOOP As Long = &H8
Public Const SND_NOSTOP As Long = &H10
Public CurrentSong As String

Public Sub PlayMidi(Song As String)
Dim i As Long

If ReadINI("CONFIG", "Music", App.Path & "\Config\Account.ini") = 1 Then
            If CurrentSong <> Song Then
                Call StopMidi
                CurrentSong = Song
                If Not Right$(Song, 4) = ".mid" Then Call PlayMP3(Song)
            End If
    If Right$(Song, 4) = ".mid" Then
        i = mciSendString("close all", 0, 0, 0)
        i = mciSendString("open """ & App.Path & "\Music\" & Song & """ Type sequencer Alias background", 0, 0, 0)
        i = mciSendString("play background notify", 0, 0, frmMirage.hwnd)
    End If
           
Else
    Call StopMidi
End If

End Sub
Public Sub StopMidi()
Dim i As Long

    If Right$(CurrentSong, 4) = ".mid" Then
        CurrentSong = vbNullString
        i = mciSendString("close all", 0, 0, 0)
    Else
        CurrentSong = vbNullString
        Call StopMP3
    End If
End Sub

Public Sub MakeMidiLoop()
Dim SBuffer As String * 256

If Right$(CurrentSong, 4) = ".mid" Then
Call mciSendString("STATUS background MODE", SBuffer, 256, 0)

If Left$(SBuffer, 7) = "stopped" Then Call mciSendString("PLAY background FROM 0", vbNullString, 0, 0)
End If

End Sub

Public Sub PlaySound(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\Config\Account.ini") = 1 Then
        If Not FileExiste("SFX\" & Sound) Then Exit Sub
        Call sndPlaySound(App.Path & "\SFX\" & Sound, SND_ASYNC Or SND_NODEFAULT)
    End If
End Sub

Public Sub StopSound()
    Dim x As Long
    Dim wFlags As Long

    wFlags = SND_ASYNC Or SND_NODEFAULT
    x = sndPlaySound("", wFlags)
End Sub

Public Sub PlayMP3(Sound As String)

'Vérifie s'il existe
If Not FileExiste("Music\" & Sound) Then Exit Sub
       
'Joue la musique
frmMirage.Mediaplayer.URL = App.Path & "\Music\" & Sound
End Sub
    
Public Sub StopMP3()
frmMirage.Mediaplayer.Close
End Sub

