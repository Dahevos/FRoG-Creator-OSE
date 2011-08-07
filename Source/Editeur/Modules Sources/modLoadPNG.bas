Attribute VB_Name = "modloadPNG"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Function LoadPNG(filename As String) As StdPicture
    On Error GoTo er:
    Dim PictureDecoder As New PAINTXLib.PictureDecoder
    Set LoadPNG = PictureDecoder.LoadPicture(filename)
    Exit Function
er:
    MsgBox "Erreur de chargement de " & filename & vbCrLf & "Verifiez qu'il soit présent.(Détails : " & Err.Number & " : " & Err.description & ")"
End Function

Public Function LoadImage(filename As String, DDraw As DirectDraw7, SDesc As DDSURFACEDESC2) As DirectDrawSurface7
    Dim TPict As StdPicture
    Set TPict = LoadPNG(filename)
    
    SDesc.lHeight = CLng((TPict.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    SDesc.lWidth = CLng((TPict.Width * 0.001) * 567 / Screen.TwipsPerPixelX)
    
    Set LoadImage = DDraw.CreateSurface(SDesc)
    
    Dim SDC As Long, TDC As Long
    SDC = LoadImage.GetDC
    TDC = CreateCompatibleDC(0)
    SelectObject TDC, TPict.Handle
    
    BitBlt SDC, 0, 0, SDesc.lWidth, SDesc.lHeight, TDC, 0, 0, vbSrcCopy
        
    LoadImage.ReleaseDC SDC
    DeleteDC TDC
       
    Set TPict = Nothing
End Function

Public Function LoadImageStretch(filename As String, Height As Long, Width As Long, DDraw As DirectDraw7, SDesc As DDSURFACEDESC2) As DirectDrawSurface7
    Dim TPict As New StdPicture
    Set TPict = LoadPNG(filename)
    
    SDesc.lHeight = Height
    SDesc.lWidth = Width
    
    Set LoadImageStretch = DDraw.CreateSurface(SDesc)
    
    Dim SDC As Long, TDC As Long
    SDC = LoadImageStretch.GetDC
    TDC = CreateCompatibleDC(0)
    SelectObject TDC, TPict.Handle
    
    StretchBlt SDC, 0, 0, Width, Height, TDC, 0, 0, CLng((TPict.Width * 0.001) * 567 / Screen.TwipsPerPixelX), CLng((TPict.Height * 0.001) * 567 / Screen.TwipsPerPixelY), vbSrcCopy
    
    LoadImageStretch.ReleaseDC SDC
    DeleteDC TDC
        
    Set TPict = Nothing
End Function
