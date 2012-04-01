Attribute VB_Name = "fullscreen"
Public Const CCDEVICENAME As Byte = 32
Public Const CCFORMNAME As Byte = 32
Public Const DISP_CHANGE_SUCCESSFUL As Byte = 0
Public Const DISP_CHANGE_RESTART As Byte = 1
Public Const DISP_CHANGE_FAILED As Integer = -1
Public Const DISP_CHANGE_BADMODE As Integer = -2
Public Const DISP_CHANGE_NOTUPDATED As Integer = -3
Public Const DISP_CHANGE_BADFLAGS As Integer = -4
Public Const DISP_CHANGE_BADPARAM As Integer = -5
Public Const CDS_UPDATEREGISTRY As Long = &H1
Public Const CDS_TEST As Long = &H2
Public Const DM_BITSPERPEL As Long = &H40000
Public Const DM_PELSWIDTH As Long = &H80000
Public Const DM_PELSHEIGHT As Long = &H100000

Public Type DEVMODE
dmDeviceName As String * CCDEVICENAME
dmSpecVersion As Integer
dmDriverVersion As Integer
dmSize As Integer
dmDriverExtra As Integer
dmFields As Long
dmOrientation As Integer
dmPaperSize As Integer
dmPaperLength As Integer
dmPaperWidth As Integer
dmScale As Integer
dmCopies As Integer
dmDefaultSource As Integer
dmPrintQuality As Integer
dmColor As Integer
dmDuplex As Integer
dmYResolution As Integer
dmTTOption As Integer
dmCollate As Integer
dmFormName As String * CCFORMNAME
dmUnusedPadding As Integer
dmBitsPerPel As Integer
dmPelsWidth As Long
dmPelsHeight As Long
dmDisplayFlags As Long
dmDisplayFrequency As Long
End Type

Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" _
(ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Declare Function ChangeDisplaySettings Lib "user32" Alias _
"ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long



