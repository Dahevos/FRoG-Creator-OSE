Attribute VB_Name = "modText"
Option Explicit

Public Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Public Const Quote As String = """"

Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = Green
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = White
Public Const EmoteColor As Byte = White
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = White
Public Const WhoColor As Byte = Grey
Public Const JoinLeftColor As Byte = Grey
Public Const NpcColor As Byte = White
Public Const AlertColor As Byte = White
Public Const NewMapColor As Byte = Grey

Public TexthDC As Long
Public GameFont As Long

Public Sub SetFont(ByVal Font As String, ByVal size As Byte)
    GameFont = CreateFont(size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 1, y + 0, Text, Len(Text))
    Call TextOut(hDC, x + 0, y + 1, Text, Len(Text))
    Call TextOut(hDC, x - 1, y - 0, Text, Len(Text))
    Call TextOut(hDC, x - 0, y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub
Public Sub DrawPlayerNameText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, RGB(0, 0, 0))
    Call TextOut(hDC, x + 1, y + 0, Text, Len(Text))
    Call TextOut(hDC, x + 0, y + 1, Text, Len(Text))
    Call TextOut(hDC, x - 1, y - 0, Text, Len(Text))
    Call TextOut(hDC, x - 0, y - 1, Text, Len(Text))
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub DrawTextInter(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String)
    Call SelectObject(hDC, GameFont)
    Call SetBkMode(hDC, vbTransparent)
    Call SetTextColor(hDC, vbBlack)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long)
Dim s As String
Dim C As Long
Dim t As Long
Dim i As Long
Dim z As Long
On Error Resume Next
t = 0
       For i = 1 To MAX_BLT_LINE
            If t = 0 Then
                If BattlePMsg(i).Index <= 0 Then
                    BattlePMsg(i).Index = 1
                    BattlePMsg(i).Msg = Msg
                    BattlePMsg(i).Color = Color
                    BattlePMsg(i).Time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).y = BattlePMsg(i).y - 15
                End If
            Else
                If BattleMMsg(i).Index <= 0 Then
                    BattleMMsg(i).Index = 1
                    BattleMMsg(i).Msg = Msg
                    BattleMMsg(i).Color = Color
                    BattleMMsg(i).Time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).y = BattleMMsg(i).y - 15
                End If
            End If
        Next i
        
        z = 1
        If t = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then If BattlePMsg(i).y < BattlePMsg(i + 1).y Then z = i Else If BattlePMsg(i).y < BattlePMsg(1).y Then z = i
            Next i
                        
            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = Msg
            BattlePMsg(z).Color = Color
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then If BattleMMsg(i).y < BattleMMsg(i + 1).y Then z = i Else If BattleMMsg(i).y < BattleMMsg(1).y Then z = i
            Next i
                        
            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = Msg
            BattleMMsg(z).Color = Color
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
End Sub


Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
    If NewLine Then Txt.Text = Txt.Text + Msg + vbCrLf Else Txt.Text = Txt.Text + Msg
    
    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Function Parse(ByVal num As Long, ByVal data As String)
Dim i As Long
Dim n As Long
Dim sChar As Long
Dim eChar As Long

    n = 0
    sChar = 1
    
    For i = 1 To Len(data)
        If Mid$(data, i, 1) = SEP_CHAR Then
            If n = num Then
                eChar = i
                Parse = Mid$(data, sChar, eChar - sChar)
                Exit For
            End If
            
            sChar = i + 1
            n = n + 1
        End If
    Next i
    
End Function

