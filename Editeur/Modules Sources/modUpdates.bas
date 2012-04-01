Attribute VB_Name = "modUpdate"
Option Explicit

Type FileData
    name As String
    version As String
    Path As String
End Type

'URL de mise à jour auto de votre éditeur
'
'à la racine de cette url doit se trouver
'le fichier Éditeur.ini ainsi que le fichier
'News.txt, le programme rename.exe ainsi que
'l'éditeur (exe). Le programme rename.exe a
'pour but de supprimer l'ancien éditeur et de
'renommer le nouveau pour ne pas briser les
'racourcis
Public Const URL As String = "http://www.markall.org/MAJ-EDIT/FROG/"

Public FileBool As Boolean
Public UpdateErr As Boolean

Public Files() As FileData
Public TempFiles() As FileData

Public Path As String
Public FileName As String

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
    sSpaces = Space$(5000)
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub EcrireINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Sub Updater()
On Error Resume Next 'éd
Uexe = False

Load frmUpdate

Call SetStatus("Recherche des mises à jour...")
CurrentVersion = ReadINI("EDITEUR", "version", App.Path & "\Editeur.ini")

DownloadFile "Editeur.ini", "temp.ini"
If UpdateErr Then Exit Sub

DownloadFile "News.txt", "News.txt"
If UpdateErr Then Exit Sub

LateVersion = ReadINI("EDITEUR", "version", App.Path & "\temp.ini")

Dim ligne As String, txt As String
If Dir$(App.Path & "\News.txt") <> vbNullString Then
    Open App.Path & "\News.txt" For Input As #1
    While Not EOF(1)
    Line Input #1, ligne
    If txt = vbNullString Then txt = ligne Else txt = txt & vbNewLine & ligne
    DoEvents
    Wend
    Close #1
End If

If Trim$(txt) = vbNullString Then frmUpdate.news.Text = "Pas d'information sur les News." Else frmUpdate.news.Text = txt

If Dir$(App.Path & "\temp.ini") <> vbNullString Then Kill App.Path & "\temp.ini"
If Dir$(App.Path & "\News.txt") <> vbNullString Then Kill App.Path & "\News.txt"
If CurrentVersion < LateVersion Then
    frmUpdate.Show
    DoEvents
    Update = True
    frmUpdate.L2.Caption = "Mise à jours à la version " & LateVersion & " !"
    DownloadFile "rename.exe", "r.exe"
    If UpdateErr Then Exit Sub
    
    DownloadFile "Editeur.exe", "Edit.exe", "Téléchargement du Nouvel Éditeur..."
    If UpdateErr Then Exit Sub
    
    CurrentVersion = LateVersion
    Call WriteINI("EDITEUR", "version", LateVersion, App.Path & "\Editeur.ini")
    frmUpdate.L3.Caption = "Téléchargement terminé !"
    MsgBox "L'éditeur doit se relancer pour prendre en compte la mise a jour."
    frmUpdate.DL.EndDownload
    Call GameDestroy
End If
End Sub

Public Sub DownloadFile(srcFileName As String, targetFileName As String, Optional status As String)
With frmUpdate.DL
If Trim$(targetFileName) <> vbNullString And Trim$(srcFileName) <> vbNullString Then
FileBool = False
frmUpdate.L3.Caption = IIf(status = vbNullString, "Téléchargement de " & srcFileName & "...", status)
Call .Download(URL & srcFileName, App.Path & "\" & targetFileName)

'Tant que le fichier n'a pas fini de charger ou qu'il n'y a pas d'erreur
Do While Not (FileBool Or UpdateErr)
DoEvents
Loop
End If
End With
End Sub

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
