VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSadScript"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type define
    sVari As String
    sValue As String
End Type

Public WithEvents SControl         As ScriptControl
Attribute SControl.VB_VarHelpID = -1
Private sAllCode()      As String
Private sSubs()         As String
Private sFunctions()    As String
Public p_colSubs        As Collection
Public p_colFuncs       As Collection
Public Path             As String
Public ASub As String
Private m_oCrypt        As clsRC4

Public Property Let FilePass(sPass As String)
    m_oCrypt.Password = sPass
End Property

'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function ReadInCode(sfile As String, sModTitle As String, msc As ScriptControl, Optional bEncrypted As Boolean = False)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim sTemp As String         'Holds each line as it comes in from the file
    Dim iTemp As Integer
    Dim sTotalTemp As String    'Holds all lines
    Dim sTempCode() As String   'Temporary Array to hold Include files
    Dim iFileNum As Integer     'File Number we're working with
    Dim sDefines() As define   'Defines to go through and change later
    'Set encryption object
    Set m_oCrypt = New clsRC4
    
    DetectScriptErr = False
    
    'Unencrypt file
    If bEncrypted = True Then
        If m_oCrypt.EncryptFile(sfile, sfile & "1") = True Then
            Kill sfile
            Name sfile & "1" As sfile
        Else
            MsgBox "Unencrypt FAILED."
            Exit Function
        End If
    End If
    
    'Start Blank
    Erase sAllCode
    ReDim sDefines(0)
    iFileNum = FreeFile         'Get new file number (thats not in use)
    
    Open sfile For Input As iFileNum   'Open file
    
    Do Until EOF(iFileNum) = True                               'Loop until file is at the end
        Line Input #iFileNum, sTemp                             'Get 1 line and put in sTemp
        sTemp = Trim$(Replace(sTemp, vbTab, ""))                 'Trim string, get rid of all tabs
        If Left$(sTemp, 1) <> "#" And Trim$(sTemp) <> vbNullString Then     'If line is a comment, ignore
            sTotalTemp = sTotalTemp & sTemp & vbNewLine         'Add line to the string
        Else
            'Yeah we got include statement
            If LCase$(Left$(sTemp, 8)) = "#include" Then
                sTemp = Mid$(sTemp, InStr(sTemp, "<") + 1, Len(sTemp) - InStr(sTemp, "<") - 1)
                Path = App.Path & "\Scripts"
                sTemp = ReturnStringFromFile(Path & "\" & sTemp)
                sTotalTemp = sTemp & vbNewLine & sTotalTemp
            ElseIf LCase$(Left$(sTemp, 7)) = "#define" Then
                sTemp = Right$(sTemp, Len(sTemp) - 8)
                sDefines(UBound(sDefines)).sVari = Mid$(sTemp, 2, InStr(sTemp, "> <") - 2)
                sDefines(UBound(sDefines)).sValue = Mid$(sTemp, InStr(sTemp, "> <") + 3, Len(sTemp) - InStr(sTemp, "> <") - 3)
                ReDim Preserve sDefines(UBound(sDefines) + 1)
            End If
        End If
    Loop
    If UBound(sDefines) <> 0 Then ReDim Preserve sDefines(UBound(sDefines) - 1)
        
    Close iFileNum                                      'Close file
    
    For iTemp = 0 To UBound(sDefines)
        sTotalTemp = Replace(sTotalTemp, sDefines(iTemp).sVari, sDefines(iTemp).sValue)
    Next
    sAllCode = Split(sTotalTemp, vbNewLine)     'Use split function and put all lines into array
    
    ReDim Preserve sAllCode(UBound(sAllCode) - 1)       'Get rid of last array element (which is blank)
    'Split string into collection
    GetSubs sAllCode
    GetFunctions sAllCode
    
    'Put collection into the script control
    msc.Modules.add sModTitle
    AddSubsToCode msc, sModTitle
    AddFuncsToCode msc, sModTitle
    
    'Encrypt file
    If bEncrypted = True Then
        If m_oCrypt.EncryptFile(sfile, sfile & "1") = True Then
            Kill sfile
            Name sfile & "1" As sfile
        Else
            MsgBox "Encrypt FAILED."
        End If
    End If
    
End Function



'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function ReturnStringFromFile(sfile As String) As String
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim sTotalTemp As String    'Holds each line as it comes in from the file
    Dim iFileNum As Integer     'File Number we're working with
    Dim sTemp As String
    iFileNum = FreeFile         'Get new file number (thats not in use)
    On Local Error GoTo filenotfound
    
    Open sfile For Input As iFileNum                   'Open file
    Err.Clear
    Do Until EOF(iFileNum) = True                               'Loop until file is at the end
        Line Input #iFileNum, sTemp                             'Get 1 line and put in sTemp
        sTemp = Trim$(Replace(sTemp, vbTab, ""))                 'Trim string, get rid of all tabs
        If Left$(sTemp, 1) <> "#" And Trim$(sTemp) <> vbNullString Then     'If line is a comment, ignore
            sTotalTemp = sTotalTemp & sTemp & vbNewLine         'Add line to the string
        Else
            'Yeah we got include statement
            If LCase$(Left$(sTemp, 8)) = "#include" Then
                sTemp = Mid$(sTemp, InStr(sTemp, "<") + 1, Len(sTemp) - InStr(sTemp, "<") - 1)
                sTemp = ReturnStringFromFile(Path & "\" & sTemp)
                sTotalTemp = sTemp & vbNewLine & sTotalTemp
            End If
            
        End If
    Loop
    Close iFileNum
    ReturnStringFromFile = sTotalTemp
    Exit Function
filenotfound:
    Exit Function
End Function


'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function GetSubs(sCode() As String)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim iCount As Integer
    Dim iTemp As Integer
    Dim sTitle As String
    Dim sSub As String
    
    Set p_colSubs = New Collection
    
    For iCount = 0 To UBound(sCode)
        sSub = vbNullString
        If LCase$(Left$(sCode(iCount), 3)) = "sub" Then
            
            For iTemp = 5 To Len(sCode(iCount))
                If Mid$(sCode(iCount), iTemp, 1) = "(" Then sTitle = Mid$(sCode(iCount), 5, iTemp - 5): Exit For
            Next
            
            Do Until LCase$(sCode(iCount)) = "end sub"
                sSub = sSub & sCode(iCount) & vbNewLine
                iCount = iCount + 1
            Loop
            sSub = sSub & sCode(iCount)
            On Error Resume Next
            p_colSubs.add sSub, sTitle
            Err.Clear
        End If
    Next
End Function


'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function AddSubsToCode(mscControl As ScriptControl, sModName As String)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim iCount As Integer
    Dim sError As String
    Dim aTemp() As String
    
    On Error GoTo err1
    
    For iCount = 1 To p_colSubs.Count
         ASub = Mid$(p_colSubs(iCount), 1, InStr(1, p_colSubs(iCount), vbNewLine))
         mscControl.Modules(sModName).AddCode p_colSubs(iCount)
         
    Next
    
    Exit Function
err1:
    'aTemp = Split(p_colSubs(iCount), vbNewLine)
    'sError = sError & "ERROR:Compiling Script :: " & Err.Description & vbNewLine
    'sError = sError & "Scripting File: " & sModName & ".thraka" & vbNewLine
    'sError = sError & "Scripting Sub: " & aTemp(0) & vbNewLine & vbNewLine
    'MsgBox "WARNING: Scripting errors can cripple the existing program.", vbCritical, "Warning"
End Function

'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function GetFunctions(sCode() As String)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim iCount As Integer
    Dim iTemp As Integer
    Dim sTitle As String
    Dim sFunc As String
    
    Set p_colFuncs = New Collection
    
    For iCount = 0 To UBound(sCode)
        sFunc = vbNullString
        If LCase$(Left$(sCode(iCount), 8)) = "function" Then
            
            For iTemp = 10 To Len(sCode(iCount))
                If Mid$(sCode(iCount), iTemp, 1) = "(" Then sTitle = Mid$(sCode(iCount), 10, iTemp - 10): Exit For
            Next
            
            Do Until LCase$(sCode(iCount)) = "end function"
                sFunc = sFunc & sCode(iCount) & vbNewLine
                iCount = iCount + 1
            Loop
            sFunc = sFunc & sCode(iCount)
            On Error Resume Next
            p_colFuncs.add sFunc, sTitle
            Err.Clear
        End If
    Next

End Function
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function AddFuncsToCode(mscControl As ScriptControl, sModName As String)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim iCount As Integer
    Dim sError As String
    Dim aTemp() As String
    
    On Error GoTo err1
    
    For iCount = 1 To p_colFuncs.Count
         mscControl.Modules(sModName).AddCode p_colFuncs(iCount)
    Next
    Exit Function
    
err1:
    'aTemp = Split(p_colFuncs(iCount), vbNewLine)
    'sError = sError & "ERROR:Compiling Script :: " & Err.Description & vbNewLine
    'sError = sError & "FILE: " & sModName & ".thraka" & vbNewLine
    'sError = sError & "FUNCTION: " & aTemp(0) & vbNewLine & vbNewLine
    'sError = sError & "More Information?"
    'MsgBox "WARNING: Scripting errors can cripple the existing program.", vbCritical, "Warning"
End Function
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub Class_Initialize()
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Set SControl = New ScriptControl
    SControl.Language = "vbScript"
End Sub
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub Class_Terminate()
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Set SControl = Nothing
    Set p_colFuncs = Nothing
    Set p_colSubs = Nothing
End Sub

'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function RunCodeReturn(sModule As String, sCode As String, ParamArray abc() As Variant) As Variant
'PRIVATE BECAUSE I CANT GET IT TO WORK
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    Dim sError As String
    Dim aTemp() As String
    Dim bTemp As Variant
    
    'How the heck to pass a paramarray array to another function using a
    'paramarray array?? if ya know email me andyd@vuetura.com
    
    'On Error GoTo ScriptErr
    If IsMissing(abc) Then bTemp = SControl.Modules(sModule).Run(sCode) Else bTemp = SControl.Modules(sModule).Run(sCode, abc)
    
    RunCodeReturn = bTemp
End Function

'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Function ExecuteStatement(sModule As String, sCode As String)
'/////////////////////////////\/\/\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ASub = sCode
    
    On Error Resume Next
    SControl.Modules(sModule).ExecuteStatement sCode
    

End Function

Private Sub SControl_Error()
    Dim sError As String
    Dim msgerr As String
    Dim temps(1 To 6) As String
    
    If Not DetectScriptErr Then
    DetectScriptErr = True
    sError = sError & "ERREUR : Exécution des Scripts : " & Err.Description & vbNewLine
    sError = sError & "Plus D'Informations?"
    temps(1) = ASub
    temps(2) = SControl.Error.Description
    temps(3) = SControl.Error.Line
    temps(4) = SControl.Error.Column
    temps(5) = SControl.Error.text
    temps(6) = SControl.Error.Source
    If MsgBox(sError, vbYesNo Or vbCritical, "Scripting Error") = vbYes Then MsgBox "Description : " & temps(2) & vbNewLine & "Ligne(du sub) : " & temps(3) & vbNewLine & "Colone : " & temps(4) & vbNewLine & "Code : " & vbNewLine & "    " & temps(5) & vbNewLine & "Source : " & temps(6) & vbNewLine & " Sub : " & temps(1)
    End If
    SControl.Error.Clear
    
End Sub


Public Function EncryptFile(sfile As String, sPassword As String)
'
'   Nice function for someone who wants to encrypt a file
'

    If m_oCrypt.EncryptFile(sfile, sfile & "1", sPassword) = True Then
        Kill sfile
        Name sfile & "1" As sfile
    Else
        MsgBox "Encryption Failed"
    End If

End Function
