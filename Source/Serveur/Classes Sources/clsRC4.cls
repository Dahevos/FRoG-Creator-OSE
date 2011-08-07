VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BLOCKSIZE = 256
Dim mnsBox(0 To 255) As Integer 'S-Box
Dim mnKeep(0 To 255) As Integer
Private mstrError As String
Private mstrPassword As String
Private mnErrorNumber As Long

Public Property Get ErrorNumber() As Long
    ErrorNumber = mnErrorNumber
End Property

Public Property Let Password(ByVal vData As String)
    mstrPassword = vData
    Initialize mstrPassword
End Property

Public Property Get Password() As String
    Password = mstrPassword
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = mstrError
End Property

Public Function EncryptFile(ByVal strSource As String, ByVal strTarget As String, Optional strPassword As String) As Boolean
Dim strNameRoutine As String ' Name of routine For logging and Error routine
Dim nResult As Long
Dim inbyte As Byte
Dim nIndex As Long
Dim nSourceFile As Integer
Dim nTargetFile As Integer
Dim nSourceSize As Long
Dim nChunkSize As Integer
Dim strInput As String
Dim strOutput As String
Dim blnContinue As Boolean
    On Local Error GoTo EncryptFile_Error
    ' InitialinIndexe variables
    strNameRoutine = "EncryptFile"
    nResult = 0 ' 0 = Failure - Must change To indicate success

    If mstrPassword = vbNullString And strPassword = vbNullString Then
        mstrError = "You need To enter a password For encrypten or decrypten"
        GoTo EncryptFile_Exit
    Else
        If Len(strPassword) And strPassword <> mstrPassword Then mstrPassword = strPassword
    End If


    If Len(strSource) = 0 Or Len(strTarget) = 0 Then mstrError = "Error - Source/Target name missing": GoTo EncryptFile_Exit
    
    If Len(Dir$(strSource)) = 0 Then mstrError = "Error missing source": GoTo EncryptFile_Exit
    
    If Len(Dir$(strTarget)) Then Kill strTarget
    
    ' get the file handles
    nSourceFile = FreeFile
    nSourceSize = FileLen(strSource)
    Open strSource For Binary As nSourceFile
    nTargetFile = FreeFile
    Open strTarget For Binary As nTargetFile
    blnContinue = False ' Set this so we reset the indexes in the first call...


    Do Until nIndex >= nSourceSize


        If nIndex + BLOCKSIZE > nSourceSize Then nChunkSize = nSourceSize - nIndex Else nChunkSize = BLOCKSIZE
        nIndex = nIndex + nChunkSize
        strInput = Space$(nChunkSize) ' init For getting data
        Get #nSourceFile, , strInput
        strOutput = EnDeCrypt(strInput, blnContinue)
        Put #nTargetFile, , strOutput
        blnContinue = True ' mark it so that we Do Not reset the indexes on subsuquent calls
    Loop
    ' clean up
    Close nSourceFile
    Close nTargetFile
    nResult = True
EncryptFile_Exit:
    On Local Error GoTo 0 ' turn off error trapping
    EncryptFile = nResult
    Exit Function
    ' Error Recovery & Logging
EncryptFile_Error:
    ' Log the error and exit routine
    mnErrorNumber = Err.Number
    mstrError = Err.description & " In " & strNameRoutine
    nResult = 0 ' verify that we are Set To failure
    Resume EncryptFile_Exit
End Function

Public Function EncryptString(ByVal strSource As String, Optional strPassword As String) As String
Dim strNameRoutine As String ' Name of routine For logging and Error routine
Dim strResult As String
    On Local Error GoTo EnCryptString_Error
    ' Initialize variables
    strNameRoutine = "EnCryptString"
    strResult = vbNullString ' 0 = Failure - Must change To indicate success
    ' make sure we have the files, names and
    '     basic requirements


    If mstrPassword = vbNullString And strPassword = vbNullString Then
        mstrError = "You need To enter a password For encrypten or decrypten"
        GoTo EnCryptString_Exit
    Else

        If Len(strPassword) And strPassword <> mstrPassword Then mstrPassword = strPassword
    End If

    If Len(strSource) = 0 Then
        mstrError = "Error - Source/Target name missing"
        GoTo EnCryptString_Exit
    End If
    strResult = EnDeCrypt(strSource, False)
EnCryptString_Exit:
    On Local Error GoTo 0 ' turn off error trapping
    EncryptString = strResult
    Exit Function
    ' Error Recovery & Logging
EnCryptString_Error:
    ' Log the error and exit routine
    mnErrorNumber = Err.Number
    mstrError = Err.description & " In " & strNameRoutine
    strResult = vbNullString ' verify that we are Set To failure
    Resume EnCryptString_Exit
End Function

Private Sub Initialize(ByVal strPassword As String)
    Dim temp As Integer
    Dim nBufferIndex As Integer
    Dim nPwdIndex As Integer
    'Save Password in Byte-Array
    nPwdIndex = 0


    For nBufferIndex = 0 To 255
        nPwdIndex = nPwdIndex + 1


        If nPwdIndex > Len(strPassword) Then nPwdIndex = 1
        mnKeep(nBufferIndex) = Asc(Mid$(strPassword, nPwdIndex, 1))
    Next nBufferIndex
    'INI S-Box


    For nBufferIndex = 0 To 255
        mnsBox(nBufferIndex) = nBufferIndex
    Next nBufferIndex
    nPwdIndex = 0


    For nBufferIndex = 0 To 255
        nPwdIndex = (nPwdIndex + mnsBox(nBufferIndex) + mnKeep(nBufferIndex)) Mod 256
        ' Swap( mnsBox(i),mnsBox(j) )
        temp = mnsBox(nBufferIndex)
        mnsBox(nBufferIndex) = mnsBox(nPwdIndex)
        mnsBox(nPwdIndex) = temp
    Next nBufferIndex
End Sub

Private Function EnDeCrypt(strSourceText As String, Optional blnContinue As Boolean) As String 'Only use this routine For short texts
    Static nIndex As Integer
    Static nIndex2 As Integer ' ok it's a poor name, but it is simply the second index...
    Dim nKeyByte As Integer
    Dim byteCypher As Byte
    Dim strCipher As String
    Dim nSwap As Integer
    Dim nTextIndex As Long


    If blnContinue = False Then
        Initialize mstrPassword ' we have To re-initialize everytime because of the array shuffle
        nIndex = 0
        nIndex2 = 0
    End If


    For nTextIndex = 1 To Len(strSourceText)
        nIndex = (nIndex + 1) Mod 256
        nIndex2 = (nIndex2 + mnsBox(nIndex)) Mod 256
        ' Swap( mnsBox(nIndex),mnsBox(nIndex2) )
        '
        nSwap = mnsBox(nIndex)
        mnsBox(nIndex) = mnsBox(nIndex2)
        mnsBox(nIndex2) = nSwap
        'Generate Keybyte nKeyByte
        nKeyByte = mnsBox((mnsBox(nIndex) + mnsBox(nIndex2)) Mod 256)
        'Plaintextbyte xor Keybyte
        byteCypher = Asc(Mid$(strSourceText, nTextIndex, 1)) Xor nKeyByte
        strCipher = strCipher & Chr$(byteCypher)
    Next nTextIndex
    EnDeCrypt = strCipher
End Function

Private Function EnDeCryptSingle(bytePlain As Byte, Optional blnContinue As Boolean) As Byte 'Use this routine For really huge files
    Static nIndex As Integer
    Static nIndex2 As Integer
    Dim nSwap As Integer
    Dim nKeyByte As Integer
    Dim byteCipher As Byte


    If blnContinue = False Then
        Initialize mstrPassword ' we have To re-initialize everytime because of the array shuffle
        nIndex = 0
        nIndex2 = 0
    End If
    ' get calculation values
    nIndex = (nIndex + 1) Mod 256
    nIndex2 = (nIndex2 + mnsBox(nIndex)) Mod 256
    ' Swap( mnsBox(nIndex),mnsBox(nIndex2) )
    '
    nSwap = mnsBox(nIndex)
    mnsBox(nIndex) = mnsBox(nIndex2)
    mnsBox(nIndex2) = nSwap
    'Generate nKeyByteeybyte nKeyByte
    nKeyByte = mnsBox((mnsBox(nIndex) + mnsBox(nIndex2)) Mod 256)
    'Plaintextbyte xor nKeyByteeybyte
    byteCipher = bytePlain Xor nKeyByte
    EnDeCryptSingle = byteCipher
End Function
