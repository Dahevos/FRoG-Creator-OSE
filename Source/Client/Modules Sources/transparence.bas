Attribute VB_Name = "transparence"
' déclaration pour fonction déplacement de la fenetre
' exemple de code :
' sub form_mousedown(...)
'     ReleaseCapture
'      SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
' endsub
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Declare Function ReleaseCapture Lib "user32" () As Long ' et de la relacher

' déclaration pour fonction TransRegion

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Public Function TransRegion(frm As Form, TranslucenceLevel As Byte, Crk As Long) As Boolean

'**************************************************
'fonction: creer un form transparante et aux forme iréguliere,
'          à partir d'une image de fond de fenetre

'PARAMETERS:
' frm: la fenêtre
' TranslucenceLevel: valeur de 0 à 255 (0 = complétement transparanet, 255 = opaque)
' Crk: couleur a utilisée comme transparance totale pour créer les contours irréguliers

' EXEMPLE:
' Private Sub Form_Load()
'   TranslucentForm Me, 128, Crk
' End Sub

'RETURNS: TRUE IF SUCCESSFUL, FALSE OTHERWISE

SetWindowLong frm.hwnd, GWL_EXSTYLE, WS_EX_LAYERED
SetLayeredWindowAttributes frm.hwnd, Crk, TranslucenceLevel, &H3

TransRegion = Err.LastDllError = 0
End Function



