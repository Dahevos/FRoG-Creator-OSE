VERSION 5.00
Begin VB.UserControl Download 
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1050
   Picture         =   "DL.ctx":0000
   ScaleHeight     =   1515
   ScaleWidth      =   1050
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   840
   End
End
Attribute VB_Name = "Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Unload()
Event Error(SubName As String, ErrNum As Long, ErrDesc As String)
Event BeginDownload(URL As String)
Event FinishDownload(File As String, FileStr As String)
Event OnProgress(BytesRead As Long, BytesMax As Long, ProgressPercent As Byte)
Event OnProgressChange(BytesRead As Long, BytesMax As Long, ProgressPercent As Byte)

Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const DoNothing = 0, Wait = 1, InProgress = 2, Finish = 3
Private status As Byte, FileName As String, LastRead As Long, LastMax As Long, Timer As Long, FileStr As String, StopIt As Boolean

Private Sub PutVar(ByVal INIFile As String, INISection As String, INIKey As String, INIValue As String)
On Error GoTo Err
    INIFile = App.Path & "\" & INIFile
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
    Exit Sub
Err:
RaiseEvent Error("PutVar", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Function GetVar(ByVal INIFile As String, INISection As String, INIKey As String) As String
On Error GoTo Err
    INIFile = App.Path & "\" & INIFile
    Dim StringBuffer As String
    Dim StringBufferSize As Long
    
    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)
    
    StringBufferSize = GetPrivateProfileString(INISection, INIKey, vbNullString, StringBuffer, StringBufferSize, INIFile)
    
    If StringBufferSize > 0 Then
        GetVar = Left$(StringBuffer, StringBufferSize)
    Else
        GetVar = vbNullString
    End If
    Exit Function
Err:
RaiseEvent Error("GetVar", Err.Number, Err.description)
Err.Clear
StopIt = True
End Function

Private Sub Timer1_Timer()
On Error GoTo Err
DoEvents
'label1.Caption = GetVar("temp.ini", "VERSION", "MaxFolder") & ":" & GetVar("Config\Dossier1.Up", "FOLDER", "FolderName") 'Status
Exit Sub
Err:
RaiseEvent Error("Timer1", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo Err
DoEvents
FileStr = AsyncProp.value
status = Finish
Exit Sub
Err:
RaiseEvent Error("ReadComplete", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
On Error GoTo Err
DoEvents
If AsyncProp.BytesMax > 0 Then
status = InProgress
RaiseEvent OnProgress(AsyncProp.BytesRead, AsyncProp.BytesMax, (AsyncProp.BytesRead / AsyncProp.BytesMax) * 100)
If LastRead <> AsyncProp.BytesRead Or LastMax <> AsyncProp.BytesMax Then RaiseEvent OnProgressChange(AsyncProp.BytesRead, AsyncProp.BytesMax, (AsyncProp.BytesRead / AsyncProp.BytesMax) * 100)
LastRead = AsyncProp.BytesRead: LastMax = AsyncProp.BytesMax
End If
'Label1.Caption = AsyncProp.Status & " - " & AsyncProp.StatusCode
If StopIt Then CancelAsyncRead: RaiseEvent Unload
Exit Sub
Err:
RaiseEvent Error("ReadProgress", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_Click()
On Error GoTo Err
Call Download("http://irtus.markall.org/Patch/info.ini", App.Path & "\temp.ini")
Call Download("http://irtus.markall.org/Patch/GFX/arrows.png", App.Path & "\GFX\arrows.png")
Exit Sub
Err:
Call MsgBox(Err.Number & vbNewLine & Err.description)
RaiseEvent Error("UC_Click", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_Initialize()
On Error GoTo Err
UserControl.Width = 750
UserControl.Height = 1425
status = DoNothing
StopIt = False
Exit Sub
Err:
RaiseEvent Error("UC_Init", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Public Property Get DownloadStatus() As Byte
On Error GoTo Err
DownloadStatus = status
Exit Property
Err:
RaiseEvent Error("DownloadStatus", Err.Number, Err.description)
Err.Clear
StopIt = True
End Property

Private Function FileExist(ByVal Chemin As String) As Boolean
On Error GoTo Err
FileExist = CBool(PathFileExists(Chemin))
Exit Function
Err:
RaiseEvent Error("FileExist", Err.Number, Err.description)
Err.Clear
StopIt = True
End Function

Public Sub EndDownload()
StopIt = True
End Sub

Public Sub Download(URL As String, File As String)
On Error GoTo Err
Dim f As Integer
FileName = File: FileStr = vbNullString
If FileExist(File) Then Kill File
RaiseEvent BeginDownload(URL)
AsyncRead URL, vbAsyncTypeFile
status = Wait
Timer = GetTickCount
Do While status <> Finish And FileStr = vbNullString And Not StopIt
If status = Wait And GetTickCount > Timer + 5000 Then
f = MsgBox("Cela fait 5 Secondes que l'éditeur n'a aucune réponse, souhaiter fermer l'application ?", vbQuestion Or vbYesNoCancel, "Délais Dépassé")
If f = vbYes Then
RaiseEvent Unload
ElseIf f = vbCancel Then
StopIt = True
Else: Timer = GetTickCount
End If
End If
DoEvents
Loop
FileCopy FileStr, FileName
Kill FileStr
RaiseEvent FinishDownload(File, FileStr)
status = DoNothing
Exit Sub
Err:
RaiseEvent Error("Download", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Function GetAsyncType(URL As String) As AsyncTypeConstants
On Error GoTo Err
Dim Data() As String
Data = Split(URL, "/")
Data = Split(Data(UBound(Data)), ".")
GetAsyncType = vbAsyncTypeFile
If LCase(Data(UBound(Data))) = "png" Or LCase(Data(UBound(Data))) = "jpg" Or LCase(Data(UBound(Data))) = "gif" Or LCase(Data(UBound(Data))) = "bmp" Then GetAsyncType = vbAsyncTypeFile
Exit Function
Err:
RaiseEvent Error("GetAType", Err.Number, Err.description)
Err.Clear
StopIt = True
End Function

Private Sub UserControl_InitProperties()
On Error GoTo Err
UserControl.Width = 750
UserControl.Height = 1425
Exit Sub
Err:
RaiseEvent Error("UC_InitProps", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_Resize()
On Error GoTo Err
UserControl.Width = 750
UserControl.Height = 1425
Exit Sub
Err:
RaiseEvent Error("UC_Res", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub

Private Sub UserControl_Show()
On Error GoTo Err
UserControl.Width = 750
UserControl.Height = 1425
Exit Sub
Err:
RaiseEvent Error("UC_Show", Err.Number, Err.description)
Err.Clear
StopIt = True
End Sub
