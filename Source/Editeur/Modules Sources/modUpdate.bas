Attribute VB_Name = "modUpdate"
Option Explicit

Type FileData
    name As String
    version As String
    Path As String
End Type

Public Files() As FileData
Public TempFiles() As FileData

Public Path As String
Public filename As String

Public CurrentVersion As String
Public LateVersion As String

Private min As Long
Public Max As Long
Public tempMax As Long
Public Extension As String
Public Update As Boolean
Public Uexe As Boolean
Public Uexen As String
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String
Dim szReturn As String
    szReturn = vbNullString
    sSpaces = Space(5000)
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub EcrireINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Sub Updater()
On Error Resume Next
Max = 0
Extension = vbNullString
Uexe = False
    
Call SetStatus("Recherche des mise a jour...")
Path = frmUpdate.Inet.URL
CurrentVersion = ReadINI("EDITEUR", "version", App.Path & "\Editeur.ini")

DownloadFile "info.ini", "temp.ini"
LateVersion = ReadINI("EDITEUR", "version", App.Path & "\temp.ini")
Max = Val(ReadINI("EDITEUR", "Max", App.Path & "\temp.ini"))

If CurrentVersion < LateVersion Then
    Update = True
    frmUpdate.news.Text = GetVar(App.Path & "\temp.ini", "EDITEUR", "news")
    If Max > 0 Then ReDim Files(1 To Max)
    If Max > 0 Then ReDim TempFiles(1 To Max)
    frmUpdate.Show
    CurrentVersion = LateVersion
    frmUpdate.L2.Caption = "Vérification des versions des fichiers..."
    If Max > 0 Then DoFiles
    If Max > 0 Then CheckFile
    Call WriteINI("EDITEUR", "version", LateVersion, App.Path & "\Editeur.ini")
    frmUpdate.L2.Caption = "Mise à jours à la version " & LateVersion & "!"
    frmUpdate.L3.Caption = "Téléchargement terminé!"
End If

If Dir$(App.Path & "\temp.ini") <> vbNullString Then Kill App.Path & "\temp.ini"
End Sub

Public Function DownloadFile(srcFileName As String, targetFileName As String)
Dim Size As Long, Remaining As Long, FFile As Integer, Chunk() As Byte

If Trim$(targetFileName) <> vbNullString Then
    If Extension <> vbNullString Then
        If Mid$(Extension, 1, 1) = "\" Or Mid$(Extension, 1, 1) = "/" Then Extension = Mid$(Extension, 2, Len(Extension))
        If Mid$(Extension, Len(Extension), Len(Extension)) = "\" Or Mid$(Extension, Len(Extension), Len(Extension)) = "/" Then Extension = Mid$(Extension, 1, Len(Extension) - 1)
        If LCase$(Dir$(App.Path & "\" & Extension, vbDirectory)) <> LCase$(Extension) Then Call MkDir$(App.Path & "\" & Extension)
    Else
        targetFileName = targetFileName
    End If
        
    frmUpdate.Inet.Execute Path & srcFileName, "GET"

    Do While frmUpdate.Inet.StillExecuting
        DoEvents
        Sleep 1
    Loop
    Size = Val(frmUpdate.Inet.GetHeader("Content-Length"))
    Remaining = 0
    FFile = FreeFile
    
    If Extension <> vbNullString Then
        Open App.Path & "\" & Extension & "\" & targetFileName For Binary Access Write As #FFile
    Else
        Open App.Path & "\" & targetFileName For Binary Access Write As #FFile
    End If
    Do Until Remaining >= Size
        If Size - Remaining > 1023 Then
            Chunk = frmUpdate.Inet.GetChunk(1024, icByteArray)
            Remaining = Remaining + 1024
        Else
            Chunk = frmUpdate.Inet.GetChunk(Size - Remaining, icByteArray)
            Remaining = Size
        End If
        Put #FFile, , Chunk

        If Size > 0 Then frmUpdate.Label4.Caption = Int((Remaining / Size) * 100) & "%"
        Sleep 1
    Loop
    Close #FFile
    
    DoEvents
    frmUpdate.L3.Caption = "Téléchargement " & srcFileName & "!"
End If
End Function

Sub DoFiles()
Dim i As Long
If Dir$(App.Path & "\temp.ini") <> vbNullString Then
    For i = 1 To Max
        TempFiles(i).name = GetVar(App.Path & "\temp.ini", "FILES", "FileName" & i)
        TempFiles(i).Path = GetVar(App.Path & "\temp.ini", "FILES", "FilePath" & i)
    Next i
End If
End Sub

Sub CheckFile()
Dim a As Long
    frmUpdate.L2.Caption = vbNullString
    
    min = 0
    For a = 1 To Max
        frmUpdate.L3.Caption = "Vérification des fichiers..."
        Extension = Trim$(TempFiles(a).Path)
                        
        min = min + 1
        frmUpdate.L2.Caption = "Téléchargement de " & TempFiles(a).name & "..."
        frmUpdate.L3.Caption = "Fichier de téléchargement: " & min & " sur " & Max
        If Mid$(TempFiles(a).name, Len(TempFiles(a).name) - 3, Len(TempFiles(a).name)) = ".exe" Then Uexe = True: Uexen = TempFiles(a).name
        DownloadFile TempFiles(a).name, TempFiles(a).name
        
        Files(a).name = TempFiles(a).name
        Files(a).Path = TempFiles(a).Path
        Files(a).version = TempFiles(a).version
        Extension = vbNullString
    Next a
End Sub
