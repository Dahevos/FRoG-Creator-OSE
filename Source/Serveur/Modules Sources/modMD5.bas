Attribute VB_Name = "modMD5"
'#############################################################################
'################### FILE ORIGINALLY FOUND HERE : ############################
'################# FICHIERS TROUVES A CETTE ADRESSE : ########################
'#### http://www.totalshareware.com/ASP/detail_view.asp?application=15351 ####
'#############################################################################
'
'
' MD5.bas - wrapper for RSA MD5 DLL
'   derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm
' Functions:
'   MD5String (some string) -> MD5 digest of the given string as 32 bytes string
'   MD5File (some filename) -> MD5 digest of the file's content as a 32 bytes string
'      returns a null terminated "FILE NOT FOUND" if unable to open the
'      given filename for input
' Bugs, complaints, etc:
'   Francisco Carlos Piragibe de Almeida
'   piragibe@esquadro.com.br
' History
'       Apr, 17 1999 - fixed the null byte problem
' Contains public domain RSA C-code for MD5 digest (see MD5-original.txt file)
' The aamd532.dll DLL MUST be somewhere in your search path
'   for this to work
Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(p As String) As String
' compute MD5 digest on a given string, returning the result
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(p)
    MDStringFix p, t, r
    MD5String = r
End Function

Public Function MD5File(f As String) As String
' compute MD5 digest on o given file, returning the result
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function
