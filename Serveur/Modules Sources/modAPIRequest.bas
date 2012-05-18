Attribute VB_Name = "modAPIRequest"
Private Const STRING_SIZE = 128
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
'
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
'
Private Declare Function InternetReadFile Lib "wininet" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
'
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" _
(ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Public Function SendAPIRequest(ByVal strUrl As String) As String
Dim hOpen As Long, hFile As Long
Dim Ret As Long, sBuffer As String * 128
Dim iResult As Integer, sData As String
hOpen = InternetOpen("VB Program", 1, vbNullString, vbNullString, 0)
If hOpen = 0 Then
Exit Function
End If
hFile = InternetOpenUrl(hOpen, strUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
If hFile = 0 Then
Else
InternetReadFile hFile, sBuffer, STRING_SIZE, Ret
sData = sBuffer
Do While Ret <> 0
InternetReadFile hFile, sBuffer, STRING_SIZE, Ret
sData = sData + Mid(sBuffer, 1, Ret)
Loop
End If
InternetCloseHandle hFile
InternetCloseHandle hOpen
SendAPIRequest = sData
sData = ""
End Function

