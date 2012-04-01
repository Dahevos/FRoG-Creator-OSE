Attribute VB_Name = "modText"
Option Explicit

Public Const Quote = vbNullString

Public Const MAX_LINES = 2000

Public Const Black = 0
Public Const Blue = 1
Public Const Green = 2
Public Const Cyan = 3
Public Const Red = 4
Public Const Magenta = 5
Public Const Brown = 6
Public Const Grey = 7
Public Const DarkGrey = 8
Public Const BrightBlue = 9
Public Const BrightGreen = 10
Public Const BrightCyan = 11
Public Const BrightRed = 12
Public Const Pink = 13
Public Const Yellow = 14
Public Const White = 15

Public SayColor As Long
Public CouleurDesGuilde As Long
Public GlobalColor As Long
Public BroadcastColor As Long
Public TellColor As Long
Public EmoteColor As Long
Public AdminColor As Long
Public HelpColor As Long
Public WhoColor As Long
Public JoinLeftColor As Long
Public NpcColor As Long
Public AlertColor As Long
Public NewMapColor As Long

Public Sub TextAdd(ByVal Txt As TextBox, Msg As String, NewLine As Boolean)
Static NumLines As Long

    If NewLine Then Txt.text = Txt.text & vbCrLf & Msg Else Txt.text = Txt.text & Msg
        
    NumLines = NumLines + 1
    If NumLines >= MAX_LINES Then Txt.text = vbNullString: NumLines = 0
    
    Txt.SelStart = Len(Txt.text)
End Sub
