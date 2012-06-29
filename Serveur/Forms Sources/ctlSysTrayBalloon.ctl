VERSION 5.00
Begin VB.UserControl ctlSysTrayBalloon 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlSysTrayBalloon.ctx":0000
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   88
   ToolboxBitmap   =   "ctlSysTrayBalloon.ctx":16C2
End
Attribute VB_Name = "ctlSysTrayBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------
' Source : http://www.vbfrance.com/codes/SYSTRAY-BALLON-SEUL-CONTROLE-UTILISATEUR_50355.aspx
'-------------------------------------------------------------------------------------------
' Codes ayant servis à l'origine :
' Original de la classe SysTray
'       http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=64701&lngWId=1
' Original du control utilisateur SubClass
'       http://www.vbfrance.com/code.aspx?ID=47896
'-------------------------------------------------------------------------------------------

' Comportement: Si plusieurs UserControl sont placés sur une forme et que vous demandez l'affichage
'               de messages en même temps sur chaque UserControl, Windows buffurise les demandes et
'               affichera les messages les uns derrière les autres.

' 2009 07 25 :
'   Animation : Suppression du controle Timer et réutilisation du Timer API
'   Animation : Possibilité de faire clignoter l'icone seule (sans autre icone en alternance)
'   Animation : Suppression Property de définition de vitesse. La vitesse est fournie avec la
'               commande de Start
'   Animation : Ajout propriété Get "BlinkIsRunning"
'   SysTray   : A la disparition du message, on supprime puis réinitialise l'icone du SysTray
'               Cette manoeuvre désactivait temporairement le SubClassing et supprimait aussi
'               le clignotement en cours --> Création Sub SysTrayRestart pour ne pas toucher
'               au SubClassing
'   Général   : Les Read/WriteProperties mémorisaient les durées (Balloon et Blink), ainsi
'               que les handles des icones.
'               Les durées étant maintenant fournies avec les commandes Start, il n'est plus
'               nécessaire de les mémoriser (en mode Création), et mémoriser des Handles est
'               déconseillé --> Suppression des Read/WriteProperties
'   Général   : Renommage des fonctions
'   Général   : Renommage de l'event "Erreur" en "PgmError"
'   Général   : "Animation" remplacé par "Blink" (variables et fonctions)
'   Général   : Correction - Lors d'un Event de la Souris, repasse la main au UserControl
'               et pas à l'icône (fermeture du menu en cas de perte du focus)
'   Général   : Détection du crash de Explorer pour réaffichage de l'icone

' 2009 07 26
'   Général   : Possibilité d'utiliser plusieurs UserControl dans un même projet (identification
'               des instances par pseudo constantes indexées)
'   Général   : Quand on cliquait sur l'icone dans le SysTray, l'image du composant sur la forme
'               qui nous accueille, apparaissait. Dû au SetForground pour éviter le maintien de
'               de l 'affichage du menu contextuel quand on clique ailleurs.
'               --> Passe le Foreground au Parent au lieu du UserControl
'   Général   : Détection du crash de Explorer : Notre icone appartenant au SysTray, il était
'               impossible de recevoir des messages. Il aurait donc fallu que ce soit la forme
'               hôte qui gère cette détection. Comme je veux que le UserControl soit complètement
'               indépendant, j'ai rajouté un Timer qui surveille le changement de handle du SysTray.

' 2009 07 28
'   Général   : Mise en application du partage de mémoire (Long) pour déterminer les valeurs des
'               constantes de Timer. Méthode issue de la source de PCPT :
'               http://www.vbfrance.com/codes/PUBLIC-SHARED-SANS-MODULE-VARIABLE-SINGLETON-IDENTIFICATION-INSTANCE_50369.aspx

' 2009 08 05
'   Général   : Ajout procédure "BalloonTipShowLast" pour réafficher le dernier message
'               Cette mémorisation sert aussi à la recherche du handle de la popup du message
'   SysTray   : Il y a un problème lors du TimeOut d'un message : La fenêtre ne se referme pas. La
'               solution adoptée était de détruire et de recréer l'icône du SysTray --> Modifié pour
'               utiliser une méthode plus propre : On envoie un message à la fenêtre du message pour
'               simuler un clic utilisateur (dans ce cas, la fermeture est Ok)
'   Général   : Remplacé les Properties "IconHandle" et "BlinkIconHandle" par "IconPicture" et
'               "BlinkIconPicture" (élément fourni = Image au lieu du handle = plus souple côté client)
'               + Suppression de leurs propriétés Get
'               Nota : dans la définition de chacune de ces propriétés, le "As Image" peut être
'                      remplacé par "As Picture" sans problème

Option Explicit

' The data type for the icon in side the task bar, very simple
Private Type NOTIFYICONDATAW        ' "W" = Version Unicode
    icoSize                 As Long ' 936, et pas 940
    icoHwnd                 As Long
    icoId                   As Long
    icoFlags                As Long
    icoCallbackMessage      As Long
    icoSource               As Long
    icoTooltip(0 To 255)    As Byte ' 256 bytes Unicode = 128 caractères "VB"
    icoState                As Long
    icoStateMask            As Long
    szInfo(0 To 511)        As Byte ' 512 bytes Unicode = 256 caractères "VB"
    uTimeOutOrVersion       As Long
    szInfoTitle(0 To 127)   As Byte ' 128 bytes Unicode = 64 caractères "VB"
    dwInfoFlags             As Long
'    guidItem                As Long ' Ne pas activer, sinon la taille passe à 940 _
                                       et le "Shell_NotifyIconW" plante
End Type

' The structure that contains all the possible types of balloons
Public Enum eBalloonIconTypes
    NIIF_NONE = &H0
    NIIF_INFO = &H1
    NIIF_WARNING = &H2
    NIIF_ERROR = &H3
    NIIF_NOSOUND = &H10
End Enum

Private Enum WindowMessageConstants
    WM_ALL = -1
    WM_NULL = &H0
    WM_CREATE = &H1
    WM_DESTROY = &H2
    WM_MOVE = &H3
    WM_SIZE = &H5
    WM_ACTIVATE = &H6
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_ENABLE = &HA
    WM_SETREDRAW = &HB
    WM_SETTEXT = &HC
    WM_GETTEXT = &HD
    WM_GETTEXTLENGTH = &HE
    WM_PAINT = &HF
    WM_CLOSE = &H10
    WM_QUERYENDSESSION = &H11
    WM_QUIT = &H12
    WM_QUERYOPEN = &H13
    WM_ERASEBKGND = &H14
    WM_SYSCOLORCHANGE = &H15
    WM_ENDSESSION = &H16
    WM_SHOWWINDOW = &H18
    WM_WININICHANGE = &H1A
    WM_SETTINGCHANGE = &H1A
    WM_DEVMODECHANGE = &H1B
    WM_ACTIVATEAPP = &H1C
    WM_FONTCHANGE = &H1D
    WM_TIMECHANGE = &H1E
    WM_CANCELMODE = &H1F
    WM_SETCURSOR = &H20
    WM_MOUSEACTIVATE = &H21
    WM_CHILDACTIVATE = &H22
    WM_QUEUESYNC = &H23
    WM_GETMINMAXINFO = &H24
    WM_PAINTICON = &H26
    WM_ICONERASEBKGND = &H27
    WM_NEXTDLGCTL = &H28
    WM_SPOOLERSTATUS = &H2A
    WM_DRAWITEM = &H2B
    WM_MEASUREITEM = &H2C
    WM_DELETEITEM = &H2D
    WM_VKEYTOITEM = &H2E
    WM_CHARTOITEM = &H2F
    WM_SETFONT = &H30
    WM_GETFONT = &H31
    WM_SETHOTKEY = &H32
    WM_GETHOTKEY = &H33
    WM_QUERYDRAGICON = &H37
    WM_COMPAREITEM = &H39
    WM_GETOBJECT = &H3D
    WM_COMPACTING = &H41
    WM_WINDOWPOSCHANGING = &H46
    WM_WINDOWPOSCHANGED = &H47
    WM_POWER = &H48
    WM_COPYDATA = &H4A
    WM_CANCELJOURNAL = &H4B
    WM_NOTIFY = &H4E
    WM_INPUTLANGCHANGEREQUEST = &H50
    WM_INPUTLANGCHANGE = &H51
    WM_TCARD = &H52
    WM_HELP = &H53
    WM_USERCHANGED = &H54
    WM_NOTIFYFORMAT = &H55
    WM_CONTEXTMENU = &H7B
    WM_STYLECHANGING = &H7C
    WM_STYLECHANGED = &H7D
    WM_DISPLAYCHANGE = &H7E
    WM_GETICON = &H7F
    WM_SETICON = &H80
    WM_NCCREATE = &H81
    WM_NCDESTROY = &H82
    WM_NCCALCSIZE = &H83
    WM_NCHITTEST = &H84
    WM_NCPAINT = &H85
    WM_NCACTIVATE = &H86
    WM_GETDLGCODE = &H87
    WM_SYNCPAINT = &H88
    WM_NCMOUSEMOVE = &HA0
    WM_NCLBUTTONDOWN = &HA1
    WM_NCLBUTTONUP = &HA2
    WM_NCLBUTTONDBLCLK = &HA3
    WM_NCRBUTTONDOWN = &HA4
    WM_NCRBUTTONUP = &HA5
    WM_NCRBUTTONDBLCLK = &HA6
    WM_NCMBUTTONDOWN = &HA7
    WM_NCMBUTTONUP = &HA8
    WM_NCMBUTTONDBLCLK = &HA9
    WM_KEYFIRST = &H100
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_DEADCHAR = &H103
    WM_SYSKEYDOWN = &H104
    WM_SYSKEYUP = &H105
    WM_SYSCHAR = &H106
    WM_SYSDEADCHAR = &H107
    WM_KEYLAST = &H108
    WM_IME_STARTCOMPOSITION = &H10D
    WM_IME_ENDCOMPOSITION = &H10E
    WM_IME_COMPOSITION = &H10F
    WM_IME_KEYLAST = &H10F
    WM_INITDIALOG = &H110
    WM_COMMAND = &H111
    WM_SYSCOMMAND = &H112
    WM_TIMER = &H113
    WM_HSCROLL = &H114
    WM_VSCROLL = &H115
    WM_INITMENU = &H116
    WM_INITMENUPOPUP = &H117
    WM_MENUSELECT = &H11F
    WM_MENUCHAR = &H120
    WM_ENTERIDLE = &H121
    WM_MENURBUTTONUP = &H122
    WM_MENUDRAG = &H123
    WM_MENUGETOBJECT = &H124
    WM_UNINITMENUPOPUP = &H125
    WM_MENUCOMMAND = &H126
    WM_CTLCOLORMSGBOX = &H132
    WM_CTLCOLOREDIT = &H133
    WM_CTLCOLORLISTBOX = &H134
    WM_CTLCOLORBTN = &H135
    WM_CTLCOLORDLG = &H136
    WM_CTLCOLORSCROLLBAR = &H137
    WM_CTLCOLORSTATIC = &H138
    WM_MOUSEFIRST = &H200
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_LBUTTONDBLCLK = &H203
    WM_RBUTTONDOWN = &H204
    WM_RBUTTONUP = &H205
    WM_RBUTTONDBLCLK = &H206
    WM_MBUTTONDOWN = &H207
    WM_MBUTTONUP = &H208
    WM_MBUTTONDBLCLK = &H209
    WM_MOUSEWHEEL = &H20A
    WM_PARENTNOTIFY = &H210
    WM_ENTERMENULOOP = &H211
    WM_EXITMENULOOP = &H212
    WM_NEXTMENU = &H213
    WM_SIZING = &H214
    WM_CAPTURECHANGED = &H215
    WM_MOVING = &H216
    WM_DEVICECHANGE = &H219
    WM_MDICREATE = &H220
    WM_MDIDESTROY = &H221
    WM_MDIACTIVATE = &H222
    WM_MDIRESTORE = &H223
    WM_MDINEXT = &H224
    WM_MDIMAXIMIZE = &H225
    WM_MDITILE = &H226
    WM_MDICASCADE = &H227
    WM_MDIICONARRANGE = &H228
    WM_MDIGETACTIVE = &H229
    WM_MDISETMENU = &H230
    WM_ENTERSIZEMOVE = &H231
    WM_EXITSIZEMOVE = &H232
    WM_DROPFILES = &H233
    WM_MDIREFRESHMENU = &H234
    WM_IME_SETCONTEXT = &H281
    WM_IME_NOTIFY = &H282
    WM_IME_CONTROL = &H283
    WM_IME_COMPOSITIONFULL = &H284
    WM_IME_SELECT = &H285
    WM_IME_CHAR = &H286
    WM_IME_REQUEST = &H288
    WM_IME_KEYDOWN = &H290
    WM_IME_KEYUP = &H291
    WM_MOUSEHOVER = &H2A1
    WM_MOUSELEAVE = &H2A3
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
    WM_RENDERFORMAT = &H305
    WM_RENDERALLFORMATS = &H306
    WM_DESTROYCLIPBOARD = &H307
    WM_DRAWCLIPBOARD = &H308
    WM_PAINTCLIPBOARD = &H309
    WM_VSCROLLCLIPBOARD = &H30A
    WM_SIZECLIPBOARD = &H30B
    WM_ASKCBFORMATNAME = &H30C
    WM_CHANGECBCHAIN = &H30D
    WM_HSCROLLCLIPBOARD = &H30E
    WM_QUERYNEWPALETTE = &H30F
    WM_PALETTEISCHANGING = &H310
    WM_PALETTECHANGED = &H311
    WM_HOTKEY = &H312
    WM_PRINT = &H317
    WM_PRINTCLIENT = &H318
    WM_HANDHELDFIRST = &H358
    WM_HANDHELDLAST = &H35F
    WM_AFXFIRST = &H360
    WM_AFXLAST = &H37F
    WM_PENWINFIRST = &H380
    WM_PENWINLAST = &H38F
    WM_USER = &H400
    WM_APP = &H8000
End Enum

Private Const GWL_WNDPROC = (-4)

' Shell_notify styles
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
' The events we can extract from the balloons
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
' Constants releated to adding and removing the icon from the task tray and response level
Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_INFO = &H10
' These inform windows what action we are about to perform with the icon
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETVERSION = &H4

Private Const NOTIFYICON_VERSION = &H3

' Méthode de variable partagée en Singleton
Private Const SECTION_MAP_READ      As Long = &H4
Private Const SECTION_MAP_WRITE     As Long = &H2
Private Const FILE_MAP_READ         As Long = SECTION_MAP_READ
Private Const FILE_MAP_WRITE        As Long = SECTION_MAP_WRITE
Private Const INVALID_HANDLE_VALUE  As Long = &HFFFFFFFF
Private Const PAGE_READWRITE        As Long = &H4
' Variable partagée : Ce nom de reconnaissance est commun à tous les UC de mon type
'   (indépendant de l'application qui le supporte)
Private Const OBJECTNAME            As String = "CodesSources_ctlSysTrayBalloon"

' Message privé dédié au Tray (filtrage des évènements)
Private Const WM_USER_TRAY = WM_USER + 1

' APIs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function Shell_NotifyIconW Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATAW) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

' APIs pour la simulation du clic sur message lors d'un TimeOut
Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

' Méthode de variable partagée en Singleton (merci PCPT @ http://www.vbfrance.com/codes/PUBLIC-SHARED-SANS-MODULE-VARIABLE-SINGLETON-IDENTIFICATION-INSTANCE_50369.aspx )
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingA" (ByVal hFile As Long, ByRef lpFileMappigAttributes As Any, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function OpenFileMapping Lib "kernel32.dll" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByRef lpBaseAddress As Any) As Long

' Le code assembleur
Private mAsm(63)                As Byte
' Adresse de l'ancien CallBack
Private mOldCallBackProc        As Long
Private lShellTrayHandle        As Long  ' Mémoire du handle du Shell_TrayWnd pour test crash Explorer
Private bCrashTimerRunning      As Boolean

' Méthode de variable partagée en Singleton
Private myID                    As Long     ' Identifiant de notre instance de control (unique dans tout le système)
Private bAutoDecrement          As Boolean  ' Sera passé à False dans l'init (on veut garder un ID unique)

' Ces pseudo constantes sont initialisées dans "UserControl_ReadProperties"
' Elles identifient le composant parmi les autres composants du même type
' Cette identification est nécessaire pour le SubClassing, pour être sûr que le message
'   reçu nous est bien destiné
' "mAPP_SYSTRAY_ID" est renvoyé dans wParam, mais il n'y a que nous qui allons recevoir
'   les messages issus de notre icône, cela ne sert pas à grand chose
Private mAPP_SYSTRAY_ID         As Long
Private mAPP_TIMER_EVENT_ID_0   As Long
Private mAPP_TIMER_EVENT_ID_1   As Long
Private mAPP_TIMER_EVENT_ID_2   As Long

' Handle de l'icone à afficher dans le SysTray
Private mIconHandle             As Long

' Mémo dernier message
Private mMessageTitle           As String
Private mMessageText            As String
Private mMessageStyle           As eBalloonIconTypes

' Durée du Timer du Balloon
Private mBalloonMilliSeconds    As Long
' Etat du Timer du Balloon
Private bBalloonTmrRunning      As Boolean
Private bBallonClickForTimeout  As Boolean

' Handle de l'icone à afficher en alternance dans le SysTray (Blink)
Private mBlinkIconHandle        As Long    ' optionel
' Durée du Timer de cycle du Flash
Private mBlinkMilliSeconds      As Long
' Etat du Timer de cycle du Flash
Private bBlinkTmrRunning        As Boolean

' These are modular level variables that allow us to determine certain aspects of the icon
' and share control of the forms events
Private mIconLoaded             As Boolean
Private mIconData               As NOTIFYICONDATAW

' Les evenements retenus pour renvoie à la forme qui nous héberge
Public Event MouseMove()    ' pas beaucoup d'intérêt puisque ce c'est le Move sur notre icône uniquement
Public Event Click()
Public Event DblClick(Button As Integer)
Public Event MouseDown(Button As Integer)
Public Event MouseUp(Button As Integer)
Public Event BalloonClosed()
Public Event BalloonClicked()
Public Event BalloonShow()
Public Event BalloonTimeOut()
Public Event PgmError(Source As String, Code As Long, Description As String)
'

' ######################################################################################################################
' /.\ NE PAS DEPLACER CETTE FONCTION /.\ '
'----------------------------------------'
' Cette fonction doit rester la premiere '
' fonction "public" du module de classe  '
'----------------------------------------'
Public Function CallBackProc(ByVal hwnd As Long, _
                             ByVal uMsg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long
    
    Dim Follow As Boolean
    
'Debug.Print Time, "CallBackProc", "hwnd "; Hex(hwnd), "Msg "; Hex(uMsg), "wP "; Hex(wParam), "lP "; Hex(lParam)
    
    If mOldCallBackProc = 0 Then Exit Function
    
    ' Par defaut le controle gère l'évènement
    ' A mettre à False quand on ne veut pas propager l'évènement à l'objet original
    Follow = True
    
    Select Case uMsg
    
        Case WM_TIMER
            If wParam = mAPP_TIMER_EVENT_ID_0 Then
                ' Teste si l'explorer a crashé en regardant si son handle a été modifié
                Call CrashTimerProc
                
            ElseIf wParam = mAPP_TIMER_EVENT_ID_1 Then
                ' Timer de fin d'apparition du Balloon
                Call BalloonTimerStop
                Call BalloonTipClose
            
            ElseIf wParam = mAPP_TIMER_EVENT_ID_2 Then
                ' Timer de fin d'un cycle d'animation
                Call BlinkSwapIcons
            End If
            Follow = False
        
        Case WM_USER_TRAY
            ' Le CallBack défini dans l'objet mIconData permet de filtrer.
            ' Les évènements sont dans lParam
            Select Case lParam
                '----------------- Balloon
                ' Ballon, qui n'est pas d'Alsace, dommage
                Case NIN_BALLOONSHOW        ' 402
                    ' Un Balloon vient d'être initialisé
                    Call BalloonTimerStart
                    RaiseEvent BalloonShow

                Case NIN_BALLOONUSERCLICK   ' 405
                    ' L'utilisateur vient de cliquer sur le ballon
                    If Not bBallonClickForTimeout Then
                        ' Cas normal
                        RaiseEvent BalloonClicked
                    Else
                        ' Cas où c'est le programme qui a cliqué pour
                        '   faire disparaitre le message en fin de TimeOut
                        RaiseEvent BalloonTimeOut
                    End If

                Case NIN_BALLOONHIDE        ' 403
                    ' Fin du ballon = TimeOut (ne fonctionne pas ici)
                    Call BalloonTimerStop
                    RaiseEvent BalloonTimeOut

                Case NIN_BALLOONTIMEOUT     ' 404
                    ' Fermé par l'utilisateur
                    ' Non non, il n'y a pas d'erreur : L'event "TimeOut" apparait
                    '   quand on ferme volontairement le message
                    Call BalloonTimerStop
                    RaiseEvent BalloonClosed
                
                '----------------- Mouse
                ' Déplacement sur l'icone (bof)
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove
                
                ' Bouton gauche
                Case WM_LBUTTONDBLCLK
                    ' Passe notre container en avant plan pour être sûr que l'éventuel menu
                    '   disparaisse si la souris clique ailleurs (problème connu)
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent DblClick(1)
                Case WM_LBUTTONUP
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent Click
                    RaiseEvent MouseUp(1)
                Case WM_LBUTTONDOWN
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent MouseDown(1)
                    
                ' Bouton droit
                Case WM_RBUTTONDBLCLK
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent DblClick(2)
                Case WM_RBUTTONUP
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent MouseUp(2)
                Case WM_RBUTTONDOWN
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent MouseDown(2)
                    
                ' Bouton central
                Case WM_MBUTTONDBLCLK
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent DblClick(4)
                Case WM_MBUTTONUP
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent MouseUp(4)
                Case WM_MBUTTONDOWN
                    ' Voir explication ci-dessus
                    Call SetForegroundWindow(UserControl.Parent.hwnd)
                    RaiseEvent MouseDown(4)
                
                Case Else
                    'Debug.Print "## Event reçu non traité. uMsg "; Hex(uMsg), "wParam "; Hex(wParam), "lParam "; Hex(lParam)
            End Select
    End Select
    ' Si l'evenement doit etre transmis à la fonction CallBackProc originale
    If Follow = True Then
        CallBackProc = CallWindowProc(mOldCallBackProc, hwnd, uMsg, wParam, lParam)
    End If

End Function


' ######################################################################################################################
' UserControl

Private Sub UserControl_Initialize()
    
    '####### Code assembleur
    Dim Ofs As Long
    Dim Ptr As Long
    
    '-----  Structure pour retrouver l'adresse de la "Me.CallBackProc" (1ère procédure)
    CopyMemory Ptr, ByVal (ObjPtr(Me)), 4
    CopyMemory Ptr, ByVal (Ptr + 489 * 4), 4
    ' Crée la veritable fonction CallBackProc (à optimiser)
    Ofs = VarPtr(mAsm(0))
    MovL Ofs, &H424448B            '8B 44 24 04          mov         eax,dword ptr [esp+4]
    MovL Ofs, &H8245C8B            '8B 5C 24 08          mov         ebx,dword ptr [esp+8]
    MovL Ofs, &HC244C8B            '8B 4C 24 0C          mov         ecx,dword ptr [esp+0Ch]
    MovL Ofs, &H1024548B           '8B 54 24 10          mov         edx,dword ptr [esp+10h]
    MovB Ofs, &H68                 '68 44 33 22 11       push        Offset RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovB Ofs, &H52                 '52                   push        edx
    MovB Ofs, &H51                 '51                   push        ecx
    MovB Ofs, &H53                 '53                   push        ebx
    MovB Ofs, &H50                 '50                   push        eax
    MovB Ofs, &H68                 '68 44 33 22 11       push        ObjPtr(Me)
    MovL Ofs, ObjPtr(Me)
    MovB Ofs, &HE8                 'E8 1E 04 00 00       call        Me.CallBackProc
    MovL Ofs, Ptr - Ofs - 4
    MovB Ofs, &HA1                 'A1 20 20 40 00       mov         eax,RetVal
    MovL Ofs, VarPtr(mAsm(59))
    MovL Ofs, &H10C2               'C2 10 00             ret         10h

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    '####### Méthode de variable partagée en Singleton :
    ' Pas de décrément lors de la fermeture car ce décrément ne fait que décompter la variable
    ' Dans notre cas, on a besoin d'un ID unique qui ne sera pas réutilisable par un autre UC
    ' La variable commune est un compteur/décompteur
    ' Il est donc normal que la valeur de l'ID ne fasse que s'incrémenter au fur et à mesure
    '   de son utilisation, même entre deux fermetures du programme
    bAutoDecrement = False
    ' Définition de notre ID
    myID = SingletonIncrement(OBJECTNAME)
    ' Calcule des valeurs des constantes à utiliser :
    '   1000 : Numéro de départ
    '      4 : Nombre de "constantes" à choisir par instance
    mAPP_SYSTRAY_ID = 1000 + (4 * (myID - 1))
    mAPP_TIMER_EVENT_ID_0 = mAPP_SYSTRAY_ID + 1
    mAPP_TIMER_EVENT_ID_1 = mAPP_SYSTRAY_ID + 2
    mAPP_TIMER_EVENT_ID_2 = mAPP_SYSTRAY_ID + 3
'Debug.Print "ID : "; myID, UserControl.Ambient.DisplayName, mAPP_SYSTRAY_ID, mAPP_TIMER_EVENT_ID_0, mAPP_TIMER_EVENT_ID_1, mAPP_TIMER_EVENT_ID_2

End Sub

' Arrete le SubClassing
Private Sub UserControl_Terminate()
    ' If the icon is still in the tray, remove it
    If mIconLoaded = True Then Call SysTrayRemoveIcon
    StopSubclassing
    If bAutoDecrement Then Call SingletonDecrement(OBJECTNAME)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 1215    ' pour le comportement de l'icone
    UserControl.Height = 375    ' de notre Ctrl sur la forme
End Sub


' ######################################################################################################################
' Propriétés
'                                                             ------- IconPicture
Public Property Set IconPicture(ByVal NewValue As Image)      ' ou As Picture
    ' Mémorise le handle de cette icone
    If mIconHandle <> NewValue.Picture.Handle Then
        mIconHandle = NewValue.Picture.Handle
        mIconData.icoSource = NewValue
        If mIconLoaded Then Call SysTrayIconRefresh
    End If
End Property
'                                                             ------- BlinkIconPicture
Public Property Set BlinkIconPicture(ByVal NewValue As Image) ' ou As Picture
    mBlinkIconHandle = NewValue.Picture.Handle
End Property
'                                                             ------- Tooltip
Public Property Get Tooltip() As String
    ' Simply return the tool tip
    Tooltip = mIconData.icoTooltip
End Property
Public Property Let Tooltip(Message128 As String)
    ' Ensure the delimiter of null is kept here
    ConvertUnicodeStringToArray Message128, mIconData.icoTooltip, 256
    If mIconLoaded Then Call SysTrayIconRefresh
End Property
'                                                             ------- BlinkIsRunning
Public Property Get BlinkIsRunning() As Boolean
    BlinkIsRunning = bBlinkTmrRunning
End Property
'                                                             ------- ID (identifiant unique)
Public Property Get ID() As Long
    ID = myID
End Property


' ######################################################################################################################
' Méthodes

Public Function SysTrayAddIcon() As Boolean

    ' Ajoute notre icône dans la barre des tâches
    
    On Error GoTo ErrorHandler
    
    If mIconLoaded = False Then
        If Initialize Then
            mIconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
            ' On peut faire un Add même si elle est déjà chargée
            If Shell_NotifyIconW(NIM_ADD, mIconData) = 1 Then
                Call Shell_NotifyIconW(NIM_SETVERSION, mIconData)
            End If
            Call StartSubclassing
            ' Témoin du SysTray chargé
            mIconLoaded = True
            ' Renvoie un Ok
            SysTrayAddIcon = True
        End If
    Else
        RaiseEvent PgmError("SysTrayAddIcon", -1, "Icône standard non définie")
    End If
    Exit Function
    
ErrorHandler:
    ' Problème
    SysTrayAddIcon = False
    RaiseEvent PgmError("SysTrayAddIcon", Err.Number, Err.Description)
End Function

Public Function SysTrayRemoveIcon()
    
    ' Supprime l'icône de la barre des tâches
    
    On Error GoTo ErrorHandler
    
    If mIconLoaded Then
        ' Stoppe éventuel balloon et clignotement
        Call BalloonTimerStop
        Call BlinkTimerStop
        ' Supprime l'icône de la barre des tâches
        Call Shell_NotifyIconW(NIM_DELETE, mIconData)
        mIconLoaded = False
        Call StopSubclassing
        ' Renvoie Ok
        SysTrayRemoveIcon = True
    Else
        RaiseEvent PgmError("SysTrayRemoveIcon", -1, "SysTray non chargé")
    End If
    Exit Function
    
ErrorHandler:
    ' Problème
    SysTrayRemoveIcon = False
    RaiseEvent PgmError("SysTrayRemoveIcon", Err.Number, Err.Description)
End Function

Public Function BalloonTipShow(ByVal Title64Unicode As String, _
                               Optional ByVal Message256Unicode As String = "", _
                               Optional ByVal Style As eBalloonIconTypes = NIIF_NONE, _
                               Optional ByVal Timeout_mSec As Long = 0) As Boolean
    
    ' You must know the following in order to
    ' use this feature properly:
    
    '      If the timeout is bigger than the systems maximum then it will be brought down.
    '      (Typically, the maximum is 30 seconds)
    
    '      If the timeout is less than the systems minimum then it will be raised upwards.
    '      (Typically, the minimum is 10 seconds)
    
    On Error GoTo ErrorHandler
    
    If mIconLoaded Then
        ' Durée (si) : Timer sera déclenché dans WinProc sur NIN_BALLOONSHOW
        mBalloonMilliSeconds = Timeout_mSec
        ' Mémorise le message
        mMessageTitle = Title64Unicode
        mMessageText = Message256Unicode
        mMessageStyle = Style
        
        With mIconData
            ' Convert the title and message into an array
            ConvertUnicodeStringToArray Message256Unicode, .szInfo, 512
            ConvertUnicodeStringToArray Title64Unicode, .szInfoTitle, 128
            
            ' Store the timeout value here and the icon
            .uTimeOutOrVersion = Timeout_mSec     ' ne sert à rien, c'est le Timer qui s'en occupera
            .dwInfoFlags = Style
            .icoFlags = NIF_INFO
        End With
        
        ' Update the icon with the new information
        Shell_NotifyIconW NIM_MODIFY, mIconData
        
        ' Completed it correctly
        BalloonTipShow = True
    Else
        RaiseEvent PgmError("BalloonTipShow", -1, "SysTray non chargé")
    End If
    Exit Function

ErrorHandler:
    ' Problème
    BalloonTipShow = False
    RaiseEvent PgmError("BalloonTipShow", Err.Number, Err.Description)
End Function

Public Sub BalloonTipShowLast()
    ' Réaffiche le dernier message mémorisé
    If mMessageTitle <> "" And mMessageText <> "" And mBalloonMilliSeconds <> 0 Then
        Call BalloonTipShow(mMessageTitle, mMessageText, mMessageStyle, mBalloonMilliSeconds)
    End If
End Sub

Public Sub BalloonTipClose()

    Dim r As Long
    
    On Error GoTo ErrorHandler
    
'    ' Bon. Avec une structure NOTIFYICONDATA (sans le W final), ces instructions suppriment
'    '   bien le message, mais avec le W, ça ne marche pas, en tous les cas, dans un contrôle
'    '   utilisateur, je pense que le problème vient de là.
'    If mIconLoaded Then
'        mIconData.szInfo(0) = 0
'        mIconData.szInfoTitle(0) = 0
'        mIconData.dwInfoFlags = NIIF_NONE
'        mIconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
'        Shell_NotifyIconW NIM_MODIFY, mIconData
'    End If
    
    ' Solution de secours :
    If mIconLoaded Then
        ' La fermeture du message fonctionne quand on clique dessus.
        ' On va donc simuler un click afin de fermer le message
        ' On recherche le handle de la fenêtre de message
        r = FindMessageObject(mMessageText)
        If r > 0 Then
            ' Ce booléen permettra de savoir si le clic est dû au programme
            bBallonClickForTimeout = True
            ' Envoie le message au message
            Call SendMessage(r, WM_MBUTTONDOWN, ByVal 0&, ByVal 0&)
            Call SendMessage(r, WM_MBUTTONUP, ByVal 0&, ByVal 0&)
            bBallonClickForTimeout = False
        Else
            ' Fenêtre message non trouvée, peut être qu'elle n'est déjà plus affichée
        End If
    End If
    Exit Sub
    
ErrorHandler:
    ' Problème
    RaiseEvent PgmError("BalloonTipClose", Err.Number, Err.Description)
End Sub

Public Sub BlinkStart(ByVal CycleMilliSeconds As Long)
    ' Vitesse limite
    If CycleMilliSeconds < 200 Then CycleMilliSeconds = 200
    ' Mémorise la vitesse demandée
    mBlinkMilliSeconds = CycleMilliSeconds
    If mIconLoaded Then
        ' Arrête éventuel clignotement en cours
        If bBlinkTmrRunning Then Call BlinkTimerStop
        ' Démarre le Timer
        Call BlinkTimerStart
    Else
        RaiseEvent PgmError("BlinkStart", -1, "SysTray non démarré")
    End If
End Sub

Public Sub BlinkStop()
    ' Stoppe le Timer
    Call BlinkTimerStop
    ' Remet l'icone standard
    mIconData.icoSource = mIconHandle
    If mIconLoaded Then Call SysTrayIconRefresh
End Sub

' ######################################################################################################################
' Méthodes pour l'utilisation de Singleton

'   *- METHODE INCREMENT -*
Friend Function SingletonIncrement(ByVal sSpaceName As String) As Long
'   incrémente la valeur LONG partagée, retourne cette valeur
'   méthode à appeler une seule fois (par la classe ou le usercontrol lui-même, lors de sa création)
'   retour à conserver durant la durée de vie de l'instance (variable privée ou lecture seule)

    Static iSingleCall  As Integer
    Dim lFM             As Long
    Dim lRet            As Long
    
'   un seul appel de fonction par instance
    iSingleCall = iSingleCall + 1
    
    If iSingleCall = 1 Then
'       filemapping
        lFM = OpenFileMapping(FILE_MAP_READ, 0, sSpaceName)

        If lFM = 0 Then
'           mapping fermé = première utilisation. création du mapping
            lFM = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_READWRITE, 0&, 4&, sSpaceName)

'           écriture première valeur = 1
            If WriteMappingValue(lFM, 1&) Then SingletonIncrement = 1

'           le décrément fermera le mapping.
        Else
'           mapping ouvert, on récupère la valeur
            If ReadMappingValue(lFM, lRet) Then
'               incrémente
                lRet = lRet + 1
                SingletonIncrement = lRet
                
'               réouverture en écriture
                Call CloseHandle(lFM)
                lFM = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_READWRITE, 0&, 4&, sSpaceName)
                
'               pas de test, si on a eu quelque chose à lire c'est qu'on a pu écrire
                Call WriteMappingValue(lFM, lRet)
'                Call CloseHandle(lFM)
            End If
        End If
    Else
'       l'appel ne se fait qu'une fois par instance d'objet
'        Err.Raise vbObjectError Or vbObject, , "Utilisation incorrecte de SingletonIncrement"
        RaiseEvent PgmError("SingletonIncrement", -3, "Utilisation incorrecte de la fonction")
    End If
End Function

'   *- METHODE DECREMENT -*
Friend Function SingletonDecrement(ByVal sSpaceName As String) As Long
'   vous pouvez faire un test de variable pour vous assurer que SingletonIncrement ait bien été appelé avant
'   nb : c'est le modèle qui doit appeler la destruction

    Static iSingleCall  As Integer
    Dim lFM             As Long
    Dim lRet            As Long
    
'   un seul appel de fonction par instance
    iSingleCall = iSingleCall + 1
    
    If iSingleCall = 1 Then
'       filemapping
        lFM = OpenFileMapping(FILE_MAP_READ, 0, sSpaceName)
    
        If lFM = 0 Then
'           mapping déjà fermé = mauvaise utilisation
            Debug.Print "Le mapping est déjà fermé, vérifiez vos appels à SingletonIncrement et SingletonDecrement"
        Else
'           mapping ouvert, on récupère la valeur
            If ReadMappingValue(lFM, lRet) Then
'               decrémente
                lRet = lRet - 1
                SingletonDecrement = lRet
    
'               réouverture en écriture
                Call CloseHandle(lFM)
                lFM = CreateFileMapping(INVALID_HANDLE_VALUE, ByVal 0&, PAGE_READWRITE, 0&, 4&, sSpaceName)
                Call WriteMappingValue(lFM, lRet)
    
'               zéro, on ferme le mapping
                If lRet = 0 Then Call CloseHandle(lFM)
            End If
        End If
    Else
'       l'appel ne se fait qu'une fois par instance d'objet
'        Err.Raise vbObjectError Or vbObject, , "Utilisation incorrecte de SingletonDecrement"
        RaiseEvent PgmError("SingletonDecrement", -3, "Utilisation incorrecte de la fonction")
    End If
End Function


' ######################################################################################################################
' Fonctions internes

Private Function Initialize() As Boolean
    
    ' Initialize the icon handler and any variables that may be required by the api call

    On Error GoTo ErrorHandler
    
    ' Mode du Ctrl obligatoire si vous devez renvoyer des coordonnées
    UserControl.ScaleMode = vbPixels
    
    If mIconHandle = 0 Then GoTo ErrorHandler
    
    With mIconData
        ' Setup the flags and other settings of the icon like we normally would using the forms settings
        .icoSize = Len(mIconData)
        .icoHwnd = UserControl.hwnd
        .icoId = mAPP_SYSTRAY_ID     ' Ne sert que lorsqu'on utilise plusieurs icones
        .icoCallbackMessage = WM_USER_TRAY  ' Filtrage des messages
        .icoSource = mIconHandle
        .icoState = NIS_SHAREDICON
        ' Setup new variables to suit the balloon message
        .uTimeOutOrVersion = NOTIFYICON_VERSION
    End With
    
    ' Completed sucessfully
    Initialize = True
    Exit Function

ErrorHandler:
    ' Problème
    Initialize = False
    RaiseEvent PgmError("Initialize", Err.Number, Err.Description)
End Function

' Le subclassing démarre en redirigeant tous les messages vers la fonction CallBackProc
' Renvoie l'adresse de l'ancienne fonction
Private Function StartSubclassing() As Long
    If UserControl.Ambient.UserMode = True Then
        ' Stoppe éventuel subcalssing précédent
        Call StopSubclassing
        ' Réorientation
        mOldCallBackProc = SetWindowLong(UserControl.hwnd, GWL_WNDPROC, VarPtr(mAsm(0)))
        ' Mémorise le handle du ShellTray
        lShellTrayHandle = FindWindow("Shell_TrayWnd", vbNullString)
        ' Démarre la surveillance du crash de Explorer
        Call CrashTimerStart
'        Debug.Print "SubClassing démarré"
    End If
End Function

' Restauration de la fonction CallBackProc classique .
Private Function StopSubclassing()
    Call CrashTimerStop
    Call BalloonTimerStop
    Call BlinkTimerStop
    If mOldCallBackProc = 0 Then Exit Function
    SetWindowLong UserControl.hwnd, GWL_WNDPROC, mOldCallBackProc
'    Debug.Print "Fin de SubClassing"
    mOldCallBackProc = 0
End Function

Private Function SysTrayIconRefresh() As Boolean
    
    ' Refresh the icon in the task tray if it exists at all
    
    On Error GoTo ErrorHandler
    
    If mIconLoaded Then
        ' Thanks to Tom Pydeski for fixing this bug.
        mIconData.icoFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        ' Only bother to refresh if it actually exists
        Call Shell_NotifyIconW(NIM_MODIFY, mIconData)
        ' No errors and were done
        SysTrayIconRefresh = True
    Else
        RaiseEvent PgmError("SysTrayIconRefresh", -1, "SysTray non chargé")
    End If
    Exit Function
    
ErrorHandler:
    ' Problème
    SysTrayIconRefresh = False
    RaiseEvent PgmError("SysTrayIconRefresh", Err.Number, Err.Description)
End Function

Private Sub BalloonTimerStart()
    If UserControl.Ambient.UserMode = True Then
        If Not bBalloonTmrRunning Then
            If mBalloonMilliSeconds > 0 Then
                ' SetTimer returns the event ID we assign if it starts successfully,
                ' so this is assigned to the Boolean flag to indicate the timer is running.
                bBalloonTmrRunning = SetTimer(UserControl.hwnd, _
                                              mAPP_TIMER_EVENT_ID_1, _
                                              mBalloonMilliSeconds, _
                                              VarPtr(mAsm(0))) = mAPP_TIMER_EVENT_ID_1
            End If
        End If
    End If
End Sub

Private Sub BalloonTimerStop()
    If bBalloonTmrRunning Then
        Call KillTimer(UserControl.hwnd, mAPP_TIMER_EVENT_ID_1)
        bBalloonTmrRunning = False
    End If
End Sub

Private Sub BlinkTimerStart()
    If UserControl.Ambient.UserMode = True Then
        If Not bBlinkTmrRunning Then
            If mBlinkMilliSeconds > 0 Then
                ' SetTimer returns the event ID we assign if it starts successfully,
                ' so this is assigned to the Boolean flag to indicate the timer is running.
                bBlinkTmrRunning = SetTimer(UserControl.hwnd, _
                                            mAPP_TIMER_EVENT_ID_2, _
                                            mBlinkMilliSeconds, _
                                            VarPtr(mAsm(0))) = mAPP_TIMER_EVENT_ID_2
            End If
        End If
    End If
End Sub

Private Sub BlinkTimerStop()
    If bBlinkTmrRunning Then
        Call KillTimer(UserControl.hwnd, mAPP_TIMER_EVENT_ID_2)
        bBlinkTmrRunning = False
    End If
End Sub

Private Sub BlinkSwapIcons()
    ' Le Timer faisant clignoter l'icône vient d'arriver à terme
    Static bEtat As Boolean
    
    If bEtat = False Then
        ' Animation frame 2 (si définie)
        bEtat = True
        mIconData.icoSource = mBlinkIconHandle
        If mIconLoaded Then Call SysTrayIconRefresh
    Else
        If mIconHandle <> 0 Then
            ' Animation frame 1
            bEtat = False
            mIconData.icoSource = mIconHandle
            If mIconLoaded Then Call SysTrayIconRefresh
        Else
            Call BlinkTimerStop
            RaiseEvent PgmError("BlinkSwapIcons", -2, "Icône standard non définie")
        End If
    End If
    
End Sub

Private Sub CrashTimerStart()
    ' Toutes les 2 secondes, on vérifiera si le handle du SysTray a changé, au cas où
    '   explorer aurait crashé (voir "CrashTimerProc")
    If UserControl.Ambient.UserMode = True Then
        If Not bCrashTimerRunning Then
            ' SetTimer returns the event ID we assign if it starts successfully,
            ' so this is assigned to the Boolean flag to indicate the timer is running.
            bCrashTimerRunning = SetTimer(UserControl.hwnd, _
                                          mAPP_TIMER_EVENT_ID_0, _
                                          ByVal 2000, _
                                          VarPtr(mAsm(0))) = mAPP_TIMER_EVENT_ID_0
        End If
    End If
End Sub

Private Sub CrashTimerStop()
    If bCrashTimerRunning Then
        Call KillTimer(UserControl.hwnd, mAPP_TIMER_EVENT_ID_0)
        bCrashTimerRunning = False
    End If
End Sub

Private Sub CrashTimerProc()
    ' Lors d'un crash de Explorer, la barre des tâches est recrée.
    ' Les icones affichées dedans ne sont pas redessinées.
    ' Puisque Explorer redémarre, il aura un nouveau handle
    ' Ici, on détecte ce changement et on relance l'apparition de notre icone
    If lShellTrayHandle <> FindWindow("Shell_TrayWnd", vbNullString) Then
        If FindWindow("Shell_TrayWnd", vbNullString) <> 0 Then
            ' Le SysTray vient de changer : réinstalle notre icone
            Dim bBlinkTmrWasRunning As Boolean
            bBlinkTmrWasRunning = bBlinkTmrRunning
            Call SysTrayRemoveIcon
            Call SysTrayAddIcon
            ' Relance le timer si Blink en service
            If bBlinkTmrWasRunning Then
                Call BlinkStart(mBlinkMilliSeconds)
            End If
        End If
    End If
End Sub

'Copie un "byte"
Private Sub MovB(Ofs As Long, ByVal value As Long)
    CopyMemory ByVal Ofs, value, 1: Ofs = Ofs + 1
End Sub

'Copy un "long"
Private Sub MovL(Ofs As Long, ByVal value As Long)
    CopyMemory ByVal Ofs, value, 4: Ofs = Ofs + 4
End Sub

Private Sub ConvertUnicodeStringToArray(ByVal sString As String, _
                                        bArray() As Byte, _
                                        ByVal lMaxSize As Long)
    ' Converts a string into a byte array then transfers it to the main array and obeying
    ' any limits that have been set
    Dim Bytes() As Byte
    Dim Pointer As Long
    Dim PointerEmpty As Long
    
    If Len(sString) > 0 Then
        ' Get the string into an array of bytes so we can use it
        Bytes = sString
        For Pointer = 0 To UBound(Bytes)
            ' Store it into the next array and exit when we have reached the limit
            bArray(Pointer) = Bytes(Pointer)
            If (Pointer = (lMaxSize - 2)) Then Exit For
        Next Pointer
        For PointerEmpty = Pointer To lMaxSize - 1
            ' Fill the rest of the array with an empty character (in this case 0)
            bArray(PointerEmpty) = 0
        Next PointerEmpty
    End If
End Sub

' ######################################################################################################################
' Méthodes pour l'utilisation de Singleton

'   *- LECTURE MAPPING OUVERT -*
Private Function ReadMappingValue(ByVal hFileMappingObject As Long, ByRef lValue As Long) As Boolean
    If hFileMappingObject = 0 Then
'       mauvais param, renvoyé par le précédent CreateFileMapping
        Debug.Print "Erreur lors du CreateFileMapping"
    Else
        Dim lMVF As Long
        lMVF = MapViewOfFile(hFileMappingObject, FILE_MAP_READ, 0&, 0&, 0&)
        If lMVF = 0 Then
            Debug.Print "Erreur lors du MapViewOfFile"
        Else
'           lecture de la valeur
            Call CopyMemory(lValue, ByVal lMVF, 4&)
            
'           fermeture de la vue
            Call UnmapViewOfFile(lMVF)
            ReadMappingValue = True
        End If
    End If
End Function

'   *- ECRITURE MAPPING OUVERT -*
Private Function WriteMappingValue(ByVal hFileMappingObject As Long, ByRef lValue As Long) As Boolean
    If hFileMappingObject = 0 Then
'       mauvais param, renvoyé par le précédent CreateFileMapping
        Debug.Print "Erreur lors du CreateFileMapping"
    Else
        Dim lMVF As Long
        lMVF = MapViewOfFile(hFileMappingObject, FILE_MAP_WRITE, 0&, 0&, 0&)
        If lMVF = 0 Then
'           erreur
            Debug.Print "Erreur lors du MapViewOfFile"
        Else
'           écriture de la valeur
            Call CopyMemory(ByVal lMVF, lValue, 4&)

'           fermeture de la vue
            Call UnmapViewOfFile(lMVF)
            WriteMappingValue = True
        End If
    End If
End Function

Private Function FindMessageObject(ByVal MessageToFind As String) As Long

    ' Contrairement à ce que peut dire WinID, la classe "tooltips_class32" n'a
    '   pas pour Parent le Shell_Tray, mais le Desktop
    ' De plus, cet objet ne retourne rien quand on utilise l'API GetWindowText
    '   Il faut passer par un SendMessage
    
    Dim lHandle As Long
    Dim sTemp As String
    Dim r As Long
        
    lHandle = GetWindow(GetDesktopWindow(), GW_CHILD)
    Do While lHandle <> 0
        ' Récupère le nom de la classe de l'enfant
        sTemp = String$(256, "^")
        r = GetClassName(lHandle, sTemp, 256)
        sTemp = Left$(sTemp, r)
        If sTemp = TOOLTIPS_CLASSA Then
            ' C'est bien une ToolTips_Calss32
            ' Récupère le message (pas le titre)
            r = SendMessage(lHandle, WM_GETTEXTLENGTH, ByVal 0&, ByVal 0&)
            If r > 255 Then r = 255
            sTemp = String$(r, " ")
            Call SendMessage(lHandle, WM_GETTEXT, r, sTemp)
            ' Supprime éventuel Chr(0) final
            sTemp = Replace(sTemp, Chr$(0), "", , , vbBinaryCompare)
            ' Comme on ne récupère que 255 caractères maxi, on ne fait le test que sur cette longueur (si)
            If Len(MessageToFind) > Len(sTemp) Then MessageToFind = Left$(MessageToFind, Len(sTemp))
            ' La comparaison finale
            If sTemp = MessageToFind Then
                FindMessageObject = lHandle
                Exit Do
            End If
        End If
        ' Sinon, recherche l'enfant suivant
        lHandle = GetWindow(lHandle, GW_HWNDNEXT)
        DoEvents
    Loop
    
End Function

