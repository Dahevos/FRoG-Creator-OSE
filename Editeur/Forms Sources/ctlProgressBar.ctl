VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000040&
   CanGetFocus     =   0   'False
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2970
   ScaleHeight     =   300
   ScaleWidth      =   2970
   ToolboxBitmap   =   "ctlProgressBar.ctx":0000
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'some general window styles and messages
Private Const WS_BORDER = &H800000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WM_SIZE As Long = &H5
Private Const WM_USER = &H400

'progress bar class string
Private Const PROGRESS_CLASSA = "msctls_progress32"

'progress bar messages
Private Const PBM_SETRANGE As Long = (WM_USER + 1)
Private Const PBM_SETPOS As Long = (WM_USER + 2)
Private Const PBM_DELTAPOS As Long = (WM_USER + 3)
Private Const PBM_SETSTEP As Long = (WM_USER + 4)
Private Const PBM_STEPIT As Long = (WM_USER + 5)
Private Const PBM_SETRANGE32 As Long = (WM_USER + 6)
Private Const PBM_GETRANGE As Long = (WM_USER + 7)
Private Const PBM_GETPOS As Long = (WM_USER + 8)
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const PBM_SETBKCOLOR As Long = 8193
Private Const PBM_SETMARQUEE As Long = (WM_USER + 10)  'XP
Private Const PBM_GETSTEP As Long = WM_USER + 13  'VISTA
Private Const PBM_GETBKCOLOR As Long = WM_USER + 14  'VISTA
Private Const PBM_GETBARCOLOR As Long = WM_USER + 15  'VISTA
Private Const PBM_SETSTATE As Long = WM_USER + 16  'VISTA
Private Const PBM_GETSTATE As Long = WM_USER + 17  'VISTA

'progress bar styles
Private Const PBS_SMOOTH As Long = &H1
Private Const PBS_VERTICAL As Long = &H4
Private Const PBS_MARQUEE As Long = &H8
Private Const PBS_SMOOTHREVERSE As Long = &H10  'VISTA

'progress bar states
Private Const PBST_NORMAL As Long = &H1  'VISTA
Private Const PBST_ERROR As Long = &H2  'VISTA
Private Const PBST_PAUSED As Long = &H3  'VISTA

'progress bar structure
Private Type PPBRANGE
  iLow As Long
  iHigh As Long
End Type

'other structures
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

'other constants
Private Const RDW_UPDATENOW = &H100
Private Const RDW_INVALIDATE = &H1
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_ID = (-12)
Private Const GWL_STYLE = (-16)
Private Const GWL_USERDATA = (-21)
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000
Private Const CLR_INVALID = -1

'to get commcontrol dll version
Private Declare Function DllGetVersion Lib "comctl32" (pdvi As DLLVERSIONINFO) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Type DLLVERSIONINFO
  cbSize As Long
  dwMajor As Long
  dwMinor As Long
  dwBuildNumber As Long
  dwPlatformID As Long
End Type

'to get windows version
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

Public Enum AppearanceConstants
  ccFlat
  cc3D
End Enum
Public Enum BorderStyleConstants
  ccNone
  ccFixedSingle
End Enum
Public Enum ScrollingConstants
  ccScrollingStandard = 0
  ccScrollingSmooth = 1
End Enum
Public Enum StateConstants
  ccStateNormal = 1
  ccStateError = 2
  ccStatePaused = 3
End Enum

'property variables
Private m_Align As AlignConstants
Private m_Appearance As AppearanceConstants
Private m_BorderStyle As BorderStyleConstants
Private m_Scrolling As ScrollingConstants
Private m_Value As Long
Private m_Max As Long
Private m_Min As Long
Private m_Step As Long
Private m_Marquee As Boolean
Private m_State As StateConstants

'private vars
Private m_hModShell32 As Long
Private dwStyle As Long, dwStyleEx As Long
Attribute dwStyleEx.VB_VarUserMemId = 1073938442
Private pbHwnd As Long
Attribute pbHwnd.VB_VarUserMemId = 1073938437

'**********************************************************************************************
'* PROGRESSBAR CONTROL PROPERTIES
'**********************************************************************************************
Public Property Let Marquee(ByVal New_Value As Boolean)
  If m_Marquee = New_Value Then Exit Property
  m_Marquee = New_Value
  PropertyChanged "Marquee"
  If pbHwnd And Ambient.UserMode Then
    pvCreate
    Call SendMessageLong(pbHwnd, PBM_SETMARQUEE, New_Value, IIf(pvIsXP, 100, 30))
  End If
  If m_Marquee = False Then
    Min = m_Min
    Max = m_Max
    value = m_Value
  End If
End Property
Public Property Get Marquee() As Boolean
  Marquee = m_Marquee
End Property

Public Property Let Scrolling(ByVal New_Value As ScrollingConstants)
  If m_Scrolling = New_Value Then Exit Property
  m_Scrolling = New_Value
  If pbHwnd Then
    pvCreate
  End If
  PropertyChanged "Scrolling"
End Property
Public Property Get Scrolling() As ScrollingConstants
  Scrolling = m_Scrolling
End Property

Public Property Get hwnd() As Long
  hwnd = pbHwnd
End Property

Public Property Let Max(ByVal New_Value As Long)
  If m_Max = New_Value Then Exit Property
  If m_Min > New_Value Then
    If Ambient.UserMode Then Err.Raise 380, App.EXEName & ".ctlProgressBar"
  Else
    m_Max = New_Value
    pvSetRange
    PropertyChanged "Max"
  End If
End Property
Public Property Get Max() As Long
  Max = m_Max
End Property

Public Property Let Min(ByVal New_Value As Long)
  If m_Min = New_Value Then Exit Property
  If New_Value > m_Max Then
    If Ambient.UserMode Then Err.Raise 380, App.EXEName & ".ctlProgressBar"
  Else
    m_Min = New_Value
    pvSetRange
    PropertyChanged "Min"
  End If
End Property
Public Property Get Min() As Long
  Min = m_Min
End Property

Public Property Let value(ByVal New_Value As Long)
If New_Value > 100 Then Exit Property
  If New_Value < m_Min Or New_Value > m_Max Then
    If Ambient.UserMode Then Err.Raise 380, App.EXEName & ".ctlProgressBar"
  Else
    m_Value = New_Value
    If Ambient.UserMode Then _
       Call SendMessage(pbHwnd, PBM_SETPOS, m_Value, 0)
    PropertyChanged "Value"
  End If
End Property
Public Property Get value() As Long
Attribute value.VB_MemberFlags = "400"
  value = m_Value
End Property

Public Property Let BorderStyle(ByVal New_Value As BorderStyleConstants)
  If m_BorderStyle = New_Value Then Exit Property
  m_BorderStyle = New_Value
  pvCreate
  PropertyChanged "BorderStyle"
End Property
Public Property Get BorderStyle() As BorderStyleConstants
  BorderStyle = m_BorderStyle
End Property

Public Property Let Appearance(ByVal New_Value As AppearanceConstants)
  If m_Appearance = New_Value Then Exit Property
  m_Appearance = New_Value
  pvCreate
  PropertyChanged "Appearance"
End Property
Public Property Get Appearance() As AppearanceConstants
  Appearance = m_Appearance
End Property

Public Property Get State() As StateConstants
Attribute State.VB_MemberFlags = "400"
  State = SendMessageLong(pbHwnd, PBM_GETSTATE, 0, 0)
End Property
Public Property Let State(ByVal New_State As StateConstants)
  Call SendMessageLong(pbHwnd, PBM_SETSTATE, New_State, 0)
  m_State = New_State
End Property


'**********************************************************************************************
'* PROGRESSBAR CONTROL METHODS
'**********************************************************************************************
Public Sub Refresh()
  Dim tR As RECT
  GetWindowRect pbHwnd, tR
  OffsetRect tR, -tR.Left, -tR.Top
  RedrawWindow pbHwnd, tR, 0&, _
               RDW_UPDATENOW Or RDW_INVALIDATE Or _
               RDW_ALLCHILDREN Or RDW_ERASE
End Sub

'**********************************************************************************************
'* INTERNAL USERCONTROL THINGS
'**********************************************************************************************
Private Sub UserControl_Initialize()
  m_hModShell32 = LoadLibrary("Shell32.dll")
  m_Appearance = cc3D
  m_Max = 100
  m_Min = 0
  m_Step = 1
End Sub

Private Sub UserControl_InitProperties()
  Appearance = cc3D
  BorderStyle = ccNone
  Scrolling = False
  pvCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  pvCreate

  Scrolling = PropBag.ReadProperty("Scrolling", False)
  Appearance = PropBag.ReadProperty("Appearance", cc3D)
  BorderStyle = PropBag.ReadProperty("BorderStyle", ccNone)
  Marquee = PropBag.ReadProperty("Marquee", False)
  Max = PropBag.ReadProperty("Max", 100)
  Min = PropBag.ReadProperty("Min", 0)
  value = PropBag.ReadProperty("Value", 0)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Appearance", m_Appearance, ccFlat
  PropBag.WriteProperty "BorderStyle", m_BorderStyle, ccNone
  PropBag.WriteProperty "Marquee", m_Marquee, False
  PropBag.WriteProperty "Max", m_Max, 100
  PropBag.WriteProperty "Min", m_Min, 0
  PropBag.WriteProperty "Scrolling", m_Scrolling, False
  PropBag.WriteProperty "Value", m_Value, 0
End Sub


Private Sub UserControl_Resize()
  MoveWindow pbHwnd, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, 1
  If m_Align = UserControl.Extender.Align Then Exit Sub
  m_Align = UserControl.Extender.Align
  Call pvCreate
End Sub

Private Sub UserControl_Terminate()
  If pbHwnd Then
    Call pvDestroy
    Call FreeLibrary(m_hModShell32)
  End If
End Sub

'**********************************************************************************************
'* PRIVATE HELPER THINGS
'**********************************************************************************************
Private Function pvIsXP() As Boolean
  Dim osv As OSVERSIONINFO
  osv.dwOSVersionInfoSize = Len(osv)
  If GetVersionEx(osv) = 1 Then
    pvIsXP = (osv.dwPlatformID = VER_PLATFORM_WIN32_NT) And (osv.dwMajorVersion = 5 And osv.dwMinorVersion = 1)
  End If
End Function

Private Function pvComCtlVersion() As Long
  Dim hMod As Long
  Dim lR As Long
  Dim lptrDLLVersion As Long
  Dim tDVI As DLLVERSIONINFO

  hMod = LoadLibrary("comctl32.dll")
  If Not (hMod = 0) Then
    lR = 0  'S_OK
    lptrDLLVersion = GetProcAddress(hMod, "DllGetVersion")
    If Not (lptrDLLVersion = 0) Then
      tDVI.cbSize = Len(tDVI)
      lR = DllGetVersion(tDVI)
      If (lR = 0) Then _
         pvComCtlVersion = tDVI.dwMajor
    Else
      'If GetProcAddress failed, then the DLL is a version previous to the one
      'shipped with IE 3.x.
      pvComCtlVersion = 4
    End If
    FreeLibrary hMod
  End If
End Function

Private Sub pvSetRange()
  Dim tPR As PPBRANGE, tPA As PPBRANGE, lR As Long
  If (pbHwnd <> 0) Then
    ' try v4.70 PBM_SETRANGE32:
    SendMessageLong pbHwnd, PBM_SETRANGE32, m_Min, m_Max
    ' check whether PBM_SETRANGE32 was supported:
    tPA.iHigh = SendMessage(pbHwnd, PBM_GETRANGE, 0, tPR)
    tPA.iLow = SendMessage(pbHwnd, PBM_GETRANGE, 1, tPR)
    If (tPA.iHigh = m_Max) And (tPA.iLow = m_Min) Then
      ' ok
    Else
      ' use the original set range message:
      lR = (m_Min And &HFFFF&)
      CopyMemory VarPtr(lR) + 2, (m_Max And &HFFFF&), 2
      SendMessage pbHwnd, PBM_SETRANGE, 0, lR
    End If
  End If
End Sub

Private Sub pvSetBorder()
  Dim lStyle As Long

  If m_Appearance = ccFlat Then
    If m_BorderStyle = ccFixedSingle Then
      pvSetStyle WS_BORDER, 0
      pvSetExStyle WS_EX_CLIENTEDGE, 0
    ElseIf m_BorderStyle = ccNone Then
      pvSetExStyle WS_EX_CLIENTEDGE, 0
    End If
  ElseIf m_Appearance = cc3D Then
    If m_BorderStyle = ccFixedSingle Then
      pvSetStyle WS_BORDER, 0
    ElseIf m_BorderStyle = ccNone Then
      pvSetStyle 0, WS_BORDER
    End If
  End If
End Sub

Private Sub pvSetExStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)
  Dim lS As Long
  If Not pbHwnd = 0 Then
    lS = GetWindowLong(pbHwnd, GWL_EXSTYLE)
    lS = lS And Not lStyleNot
    lS = lS Or lStyle
    SetWindowLong pbHwnd, GWL_EXSTYLE, lS
    SetWindowPos pbHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
  End If
End Sub

Public Function pvHiWord(ByVal dwValue As Long) As Long
  Call CopyMemory(pvHiWord, ByVal VarPtr(dwValue) + 2, 2)
End Function

Public Function pvLoWord(ByVal dwValue As Long) As Long
  Call CopyMemory(pvLoWord, dwValue, 2)
End Function

Public Function pvMakeLong(ByVal wLow As Long, ByVal wHi As Long) As Long
  If (wHi And &H8000&) Then
    pvMakeLong = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
  Else
    pvMakeLong = pvLoWord(wLow) Or (&H10000 * pvLoWord(wHi))
  End If
End Function

Private Function pvClientSiteAvailable() As Boolean
  Dim a
  On Error Resume Next
  pvClientSiteAvailable = UserControl.Parent.hwnd >= 1
End Function

Private Sub pvSetStyle(ByVal lStyle As Long, ByVal lStyleNot As Long)
  Dim lS As Long
  lS = GetWindowLong(pbHwnd, GWL_STYLE)
  lS = lS And Not lStyleNot
  lS = lS Or lStyle
  Call SetWindowLong(pbHwnd, GWL_STYLE, lS)
  Call SetWindowPos(pbHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub

Private Function pvCreate() As Boolean
  'If Ambient.UserMode = False Then Exit Function
  pvDestroy
  InitCommonControls
  dwStyle = WS_VISIBLE Or WS_CHILD
  If m_Align = vbAlignLeft Or m_Align = vbAlignRight Then dwStyle = dwStyle Or PBS_VERTICAL
  If Scrolling = ccScrollingSmooth And Ambient.UserMode = True Then dwStyle = dwStyle Or PBS_SMOOTH
  If Marquee = True And Ambient.UserMode = True Then dwStyle = dwStyle Or PBS_MARQUEE

  pbHwnd = CreateWindowEx(0, PROGRESS_CLASSA, "", _
                          dwStyle, 0, 0, UserControl.ScaleWidth \ Screen.TwipsPerPixelX, UserControl.ScaleHeight \ Screen.TwipsPerPixelY, _
                          UserControl.hwnd, 0&, App.hInstance, 0&)
  UserControl.BackColor = vbButtonFace

  If pbHwnd Then
    pvSetBorder
    If Ambient.UserMode = True Then
      pvSetRange
      If m_Marquee = True Then Call SendMessageLong(pbHwnd, PBM_SETMARQUEE, True, IIf(pvIsXP, 100, 30))
      Call SendMessage(pbHwnd, PBM_SETPOS, m_Value, 0)
      If m_State <> ccStateNormal Then Call SendMessageLong(pbHwnd, PBM_SETSTATE, m_State, 0)
    Else
      Call SendMessage(pbHwnd, PBM_SETPOS, m_Max, 0)
    End If
    Refresh
    pvCreate = True
  End If
End Function

Private Sub pvDestroy()
  If pbHwnd Then
    ShowWindow pbHwnd, SW_HIDE
    SetParent pbHwnd, 0
    DestroyWindow pbHwnd
  End If
End Sub
