VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-- Class Constantes
Private Const OFFSET_4 = 4294967296#    '&H100000000
Private Const MAXINT_4 = &H7FFFFFFF

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22

Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20

Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23

Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

'-- Class Variables
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte
Private MD5_Transform '(0 To 63) As Double

'-- Class Publics Functions
'
' Function to quickly digest a file into a hex string
'
Public Function MD5FileToHexStr(sFileName As String) As String
Dim FF As Integer
    FF = FreeFile
    Open sFileName For Binary Access Read As #FF
    MD5Init
    Do While Not EOF(FF)
        Get #FF, , ByteBuffer
        If Loc(FF) < LOF(FF) Then
            ByteCounter = ByteCounter + 64
            MD5Transform ByteBuffer
        End If
    Loop
    ByteCounter = ByteCounter + (LOF(FF) Mod 64)
    Close #FF
    MD5Final
    MD5FileToHexStr = LCase(GetValues)
End Function

'
' Function to digest a text string and output the result as a string
' of hexadecimal characters.
'
Public Function MD5StrToHexStr(sSource As String) As String
    MD5Init
    MD5Update Len(sSource), StringToArray(sSource)
    MD5Final
    MD5StrToHexStr = LCase(GetValues)
End Function

'
' Initialise la classe
'   Doit être appelé avant tout calcul
'
Private Sub MD5Init()
    ByteCounter = 0
    'L'initialisation de Dobbertin
    State(1) = CLng(&H67452301)
    State(2) = CLng(&HEFCDAB89)
    State(3) = CLng(&H98BADCFE)
    State(4) = CLng(&H10325476)
    'Tableau des constantes de transformation
    MD5_Transform = Array( _
               &HD76AA478, &HE8C7B756, &H242070DB, &HC1BDCEEE, &HF57C0FAF, &H4787C62A, &HA8304613, &HFD469501, _
               &H698098D8, &H8B44F7AF, &HFFFF5BB1, &H895CD7BE, &H6B901122, &HFD987193, &HA679438E, &H49B40821, _
               &HF61E2562, &HC040B340, &H265E5A51, &HE9B6C7AA, &HD62F105D, &H2441453, &HD8A1E681, &HE7D3FBC8, _
               &H21E1CDE6, &HC33707D6, &HF4D50D87, &H455A14ED, &HA9E3E905, &HFCEFA3F8, &H676F02D9, &H8D2A4C8A, _
               &HFFFA3942, &H8771F681, &H6D9D6122, &HFDE5380C, &HA4BEEA44, &H4BDECFA9, &HF6BB4B60, &HBEBFBC70, _
               &H289B7EC6, &HEAA127FA, &HD4EF3085, &H4881D05, &HD9D4D039, &HE6DB99E5, &H1FA27CF8, &HC4AC5665, _
               &HF4292244, &H432AFF97, &HAB9423A7, &HFC93A039, &H655B59C3, &H8F0CCC92, &HFFEFF47D, &H85845DD1, _
               &H6FA87E4F, &HFE2CE6E0, &HA3014314, &H4E0811A1, &HF7537E82, &HBD3AF235, &H2AD7D2BB, &HEB86D391)
End Sub

'
' MD5 Transformation
'
Private Sub MD5Transform(Buffer() As Byte)
Dim x(16) As Long
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long

    a = State(1)
    b = State(2)
    c = State(3)
    d = State(4)

    Decode 64, x, Buffer

    ' Round 1
    FF a, b, c, d, x(0), S11, MD5_Transform(0)
    FF d, a, b, c, x(1), S12, MD5_Transform(1)
    FF c, d, a, b, x(2), S13, MD5_Transform(2)
    FF b, c, d, a, x(3), S14, MD5_Transform(3)
    FF a, b, c, d, x(4), S11, MD5_Transform(4)
    FF d, a, b, c, x(5), S12, MD5_Transform(5)
    FF c, d, a, b, x(6), S13, MD5_Transform(6)
    FF b, c, d, a, x(7), S14, MD5_Transform(7)
    FF a, b, c, d, x(8), S11, MD5_Transform(8)
    FF d, a, b, c, x(9), S12, MD5_Transform(9)
    FF c, d, a, b, x(10), S13, MD5_Transform(10)
    FF b, c, d, a, x(11), S14, MD5_Transform(11)
    FF a, b, c, d, x(12), S11, MD5_Transform(12)
    FF d, a, b, c, x(13), S12, MD5_Transform(13)
    FF c, d, a, b, x(14), S13, MD5_Transform(14)
    FF b, c, d, a, x(15), S14, MD5_Transform(15)

    ' Round 2
    GG a, b, c, d, x(1), S21, MD5_Transform(16)
    GG d, a, b, c, x(6), S22, MD5_Transform(17)
    GG c, d, a, b, x(11), S23, MD5_Transform(18)
    GG b, c, d, a, x(0), S24, MD5_Transform(19)
    GG a, b, c, d, x(5), S21, MD5_Transform(20)
    GG d, a, b, c, x(10), S22, MD5_Transform(21)
    GG c, d, a, b, x(15), S23, MD5_Transform(22)
    GG b, c, d, a, x(4), S24, MD5_Transform(23)
    GG a, b, c, d, x(9), S21, MD5_Transform(24)
    GG d, a, b, c, x(14), S22, MD5_Transform(25)
    GG c, d, a, b, x(3), S23, MD5_Transform(26)
    GG b, c, d, a, x(8), S24, MD5_Transform(27)
    GG a, b, c, d, x(13), S21, MD5_Transform(28)
    GG d, a, b, c, x(2), S22, MD5_Transform(29)
    GG c, d, a, b, x(7), S23, MD5_Transform(30)
    GG b, c, d, a, x(12), S24, MD5_Transform(31)

    ' Round 3
    HH a, b, c, d, x(5), S31, MD5_Transform(32)
    HH d, a, b, c, x(8), S32, MD5_Transform(33)
    HH c, d, a, b, x(11), S33, MD5_Transform(34)
    HH b, c, d, a, x(14), S34, MD5_Transform(35)
    HH a, b, c, d, x(1), S31, MD5_Transform(36)
    HH d, a, b, c, x(4), S32, MD5_Transform(37)
    HH c, d, a, b, x(7), S33, MD5_Transform(38)
    HH b, c, d, a, x(10), S34, MD5_Transform(39)
    HH a, b, c, d, x(13), S31, MD5_Transform(40)
    HH d, a, b, c, x(0), S32, MD5_Transform(41)
    HH c, d, a, b, x(3), S33, MD5_Transform(42)
    HH b, c, d, a, x(6), S34, MD5_Transform(43)
    HH a, b, c, d, x(9), S31, MD5_Transform(44)
    HH d, a, b, c, x(12), S32, MD5_Transform(45)
    HH c, d, a, b, x(15), S33, MD5_Transform(46)
    HH b, c, d, a, x(2), S34, MD5_Transform(47)

    ' Round 4
    II a, b, c, d, x(0), S41, MD5_Transform(48)
    II d, a, b, c, x(7), S42, MD5_Transform(49)
    II c, d, a, b, x(14), S43, MD5_Transform(50)
    II b, c, d, a, x(5), S44, MD5_Transform(51)
    II a, b, c, d, x(12), S41, MD5_Transform(52)
    II d, a, b, c, x(3), S42, MD5_Transform(53)
    II c, d, a, b, x(10), S43, MD5_Transform(54)
    II b, c, d, a, x(1), S44, MD5_Transform(55)
    II a, b, c, d, x(8), S41, MD5_Transform(56)
    II d, a, b, c, x(15), S42, MD5_Transform(57)
    II c, d, a, b, x(6), S43, MD5_Transform(58)
    II b, c, d, a, x(13), S44, MD5_Transform(59)
    II a, b, c, d, x(4), S41, MD5_Transform(60)
    II d, a, b, c, x(11), S42, MD5_Transform(61)
    II c, d, a, b, x(2), S43, MD5_Transform(62)
    II b, c, d, a, x(9), S44, MD5_Transform(63)


    State(1) = LongOverflowAdd(State(1), a)
    State(2) = LongOverflowAdd(State(2), b)
    State(3) = LongOverflowAdd(State(3), c)
    State(4) = LongOverflowAdd(State(4), d)

End Sub

'
'-- Break up input stream into 64 byte chunks
'
Private Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
Dim II As Integer
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim lngBufferedBytes As Long
Dim lngBufferRemaining As Long
Dim lngRem As Long

    lngBufferedBytes = ByteCounter Mod 64
    lngBufferRemaining = 64 - lngBufferedBytes
    ByteCounter = ByteCounter + InputLen
    ' Use up old buffer results first
    If InputLen >= lngBufferRemaining Then
        For II = 0 To lngBufferRemaining - 1
            ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
        Next II
        MD5Transform ByteBuffer

        lngRem = (InputLen) Mod 64
        ' The transfer is a multiple of 64 lets do some transformations
        For i = lngBufferRemaining To InputLen - II - lngRem Step 64
            For J = 0 To 63
                ByteBuffer(J) = InputBuffer(i + J)
            Next J
            MD5Transform ByteBuffer
        Next i
        lngBufferedBytes = 0
    Else
        i = 0
    End If

    ' Buffer any remaining input
    For K = 0 To InputLen - i - 1
        ByteBuffer(lngBufferedBytes + K) = InputBuffer(i + K)
    Next K

End Sub

'
' MD5 Finalisation
'
Private Sub MD5Final()
Dim dblBits As Double

Dim padding(72) As Byte
Dim lngBytesBuffered As Long

    padding(0) = &H80

    dblBits = ByteCounter * 8

    ' Pad out
    lngBytesBuffered = ByteCounter Mod 64
    If lngBytesBuffered <= 56 Then
        MD5Update 56 - lngBytesBuffered, padding
    Else
        MD5Update 120 - ByteCounter, padding
    End If


    padding(0) = UnsignedToLong(dblBits) And &HFF&
    padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
    padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
    padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
    padding(4) = 0
    padding(5) = 0
    padding(6) = 0
    padding(7) = 0

    MD5Update 8, padding
End Sub

'-- Fonction utilitaire Concatenation des 4 valeurs d'état en une string
Private Function GetValues() As String
    GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function

'-- Class Privates Functions

'-- Fonction utilitaire Convertir une string en array of bytes.
Private Function StringToArray(InString As String) As Byte()
Dim i As Integer
Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(InString))
    For i = 0 To Len(InString) - 1
        bytBuffer(i) = Asc(Mid(InString, i + 1, 1))
    Next i
    StringToArray = bytBuffer
End Function

'-- Fonction utilitaire Conversion d'un Long en une Hex string
Private Function LongToString(Num As Long) As String
Dim a As Byte
Dim b As Byte
Dim c As Byte
Dim d As Byte

    a = Num And &HFF&
    If a < 16 Then
        LongToString = "0" & Hex(a)
    Else
        LongToString = Hex(a)
    End If

    b = (Num And &HFF00&) \ 256
    If b < 16 Then
        LongToString = LongToString & "0" & Hex(b)
    Else
        LongToString = LongToString & Hex(b)
    End If

    c = (Num And &HFF0000) \ 65536
    If c < 16 Then
        LongToString = LongToString & "0" & Hex(c)
    Else
        LongToString = LongToString & Hex(c)
    End If

    If Num < 0 Then
        d = ((Num And &H7F000000) \ 16777216) Or &H80&
    Else
        d = (Num And &HFF000000) \ 16777216
    End If

    If d < 16 Then
        LongToString = LongToString & "0" & Hex(d)
    Else
        LongToString = LongToString & Hex(d)
    End If

End Function

'-- Fonction utilitaire
Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
Dim intDblIndex As Integer
Dim intByteIndex As Integer
Dim dblSum As Double

    intDblIndex = 0
    For intByteIndex = 0 To Length - 1 Step 4
        dblSum = InputBuffer(intByteIndex) + _
                 InputBuffer(intByteIndex + 1) * 256# + _
                 InputBuffer(intByteIndex + 2) * 65536# + _
                 InputBuffer(intByteIndex + 3) * 16777216#
        OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
        intDblIndex = intDblIndex + 1
    Next intByteIndex
End Sub

'-- FF, GG, HH, et II : Transformations pour rounds 1, 2, 3, et 4.
'     La rotation est séparée de l'addition pour empêcher un recalcul.

Private Function FF(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ByVal ac As Long) As Long
    a = LongOverflowAdd4(a, (b And c) Or (Not (b) And d), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function GG(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ByVal ac As Long) As Long
    a = LongOverflowAdd4(a, (b And d) Or (c And Not (d)), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function HH(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ByVal ac As Long) As Long
    a = LongOverflowAdd4(a, b Xor c Xor d, x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

Private Function II(a As Long, b As Long, c As Long, d As Long, x As Long, s As Long, ByVal ac As Long) As Long
    a = LongOverflowAdd4(a, c Xor (b Or Not (d)), x, ac)
    a = LongLeftRotate(a, s)
    a = LongOverflowAdd(a, b)
End Function

'-- Fonction utilitaire Rotation a gauche d'un entier long
Private Function LongLeftRotate(value As Long, bits As Long) As Long
Dim lngSign As Long
Dim lngI As Long
    bits = bits Mod 32
    If bits = 0 Then LongLeftRotate = value: Exit Function
    For lngI = 1 To bits
        lngSign = value And &HC0000000
        value = (value And &H3FFFFFFF) * 2
        value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next
    LongLeftRotate = value
End Function

'-- Fonction utilitaire Additioner 2 nombres non signés
'      Les dépassements (Overflows) sont ignorés.
Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
Dim lngHighWord As Long
Dim lngLowWord As Long
Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

'-- Fonction utilitaire Additioner 4 nombres non signés
'      Les dépassements (Overflows) sont ignorés.
Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
Dim lngHighWord As Long
Dim lngLowWord As Long
Dim lngOverflow As Long

    lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + _
                   ((Val2 And &HFFFF0000) \ 65536) + _
                   ((val3 And &HFFFF0000) \ 65536) + _
                   ((val4 And &HFFFF0000) \ 65536) + _
                   lngOverflow) And &HFFFF&
    LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function

'-- Fonction utilitaire Conversion d'un double non signé en entier long
Private Function UnsignedToLong(value As Double) As Long
If value < 0 Or value >= OFFSET_4 Then Error 6    ' Overflow
    If value <= MAXINT_4 Then
        UnsignedToLong = value
    Else
        UnsignedToLong = value - OFFSET_4
    End If
End Function

'-- Fonction utilitaire Conversion d'un entier long en double non signé
Private Function LongToUnsigned(value As Long) As Double
    If value < 0 Then
        LongToUnsigned = value + OFFSET_4
    Else
        LongToUnsigned = value
    End If
End Function



