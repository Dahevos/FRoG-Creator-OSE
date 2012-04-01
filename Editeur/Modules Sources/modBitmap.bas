Attribute VB_Name = "modBitmap"
Option Explicit

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const LR_LOADFROMFILE = &H10
Public Const IMAGE_BITMAP = 0
Public Const LR_CREATEDIBSECTION = &H2000

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

Public Type BITMAP '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Type GAME_BITMAP
    hBitmap As Long
    bmp As BITMAP
    hdc As Long
End Type

Public Sub CreateGameBitmap(ByVal FileName As String, GameBitmap As GAME_BITMAP)
    With GameBitmap
        ' Load bitmap from file
        .hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
                
        ' Get the bitmap header info
        Call GetObject(.hBitmap, Len(.bmp), .bmp)
            
        ' Holds the bitmap itself
        .hdc = CreateCompatibleDC(ByVal 0)
            
        ' Puts the image into the DC
        Call SelectObject(.hdc, .hBitmap)
    End With
End Sub

Public Sub CreateGameBackBuffer(ByVal Width As Long, ByVal Height As Long, GameBitmap As GAME_BITMAP)
    With GameBitmap
        ' Load bitmap from file
        .hBitmap = CreateCompatibleBitmap(GetDC(0), Width, Height)
                
        ' Get the bitmap header info
        Call GetObject(.hBitmap, Len(.bmp), .bmp)
            
        ' Holds the bitmap itself
        .hdc = CreateCompatibleDC(ByVal 0)
            
        ' Puts the image into the DC
        Call SelectObject(.hdc, .hBitmap)
    End With
End Sub

Public Sub DestroyGameBitmap(GameBitmap As GAME_BITMAP)
    With GameBitmap
        Call DeleteDC(.hdc)
        Call DeleteObject(.hBitmap)
    End With
End Sub

