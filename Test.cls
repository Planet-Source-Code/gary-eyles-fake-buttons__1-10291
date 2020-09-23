VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_GETSYSMENU = &H313

Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type

Private Const MFT_STRING = 0
Private Const WM_MEASUREITEM = &H2C
Private Const MF_OWNERDRAW = &H100&
Private Const MF_BYPOSITION = &H400&
'Private Const GWL_WNDPROC = (-4)
'Private Const WM_DRAWITEM = &H2B
'Private Const WM_MEASUREITEM = &H2C
'Private Const WM_COMMAND = &H111
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32" _
    Alias "SetMenuItemInfoA" _
   (ByVal hMenu As Long, ByVal uItem As Long, _
    ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long


Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Private Const LR_LOADMAP3DCOLORS = &H1000
    Private Const LR_LOADFROMFILE = &H10
    Private Const LR_LOADTRANSPARENT = &H20
    Private Const IMAGE_BITMAP = 0

Private Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8


Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap%) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function PatBlt% Lib "gdi32" (ByVal hdc%, ByVal x%, ByVal y%, ByVal nWidth%, ByVal nHeight%, ByVal dwRop&)
Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Const WM_SYSCOLORCHANGE = &H15
Private Const WM_SETCURSOR = &H20
Private Const WM_CUSTOM = &H400
Private WithEvents m_t1 As CTimer
Attribute m_t1.VB_VarHelpID = -1

Public Event Resizing(hwnd As Long, focus As Boolean, resizetype As Long, syscolchange As Boolean)

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

' SetWindowsHook() codes
'Public Const WH_MIN = (-1)
'Public Const WH_MSGFILTER = (-1)
'Public Const WH_JOURNALRECORD = 0
'Public Const WH_JOURNALPLAYBACK = 1
'Public Const WH_KEYBOARD = 2
'Public Const WH_GETMESSAGE = 3
'Public Const WH_CALLWNDPROC = 4
'Public Const WH_CBT = 5
'Public Const WH_SYSMSGFILTER = 6
Private Const WH_MOUSE = 7
'Public Const WH_HARDWARE = 8
'Public Const WH_DEBUG = 9
'Public Const WH_SHELL = 10
'Public Const WH_FOREGROUNDIDLE = 11
'Public Const WH_MAX = 11

' Hook Codes
'Public Const HC_ACTION = 0
'Public Const HC_GETNEXT = 1
'Public Const HC_SKIP = 2
'Public Const HC_NOREMOVE = 3
'Public Const HC_NOREM = HC_NOREMOVE
'Public Const HC_SYSMODALON = 4
'Public Const HC_SYSMODALOFF = 5

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Const WM_PAINT = &HF
'Private Const WM_MOVE = &H3
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_COMMAND = &H111
Private Const WM_CLOSE = &H10
Private Const WM_MOVE = &H3
'Public Const WM_SIZE = &H5

'These declarations draw the caption buttons...
Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, _
    lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const DFC_CAPTION = 1

Public Enum SysButtons
    DFCS_CAPTIONRESTORE = &H3
    DFCS_CAPTIONMIN = &H1
    DFCS_CAPTIONMAX = &H2
    DFCS_CAPTIONHELP = &H4
    DFCS_CAPTIONCLOSE = &H0
End Enum

Public Enum SysButtonsI
    DFCS_DEPRESSED = &H200
    DFCS_INACTIVE = &H100
    DFCS_ACTIVE = 0
End Enum

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Const MF_SYSMENU = &H2000&

Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MIIM_SUBMENU = &H4&
Private Const MIIM_CHECKMARKS = &H8&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Const WM_MENUSELECT = &H11F
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Boolean

Private Const SC_RESTORE = &HF120&
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const HTSYSMENU = 3
Private Const WM_NCHITTEST = &H84

Private Declare Function GetSystemMenu Lib "user32" _
     (ByVal hwnd As Long, _
     ByVal bRevert As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
'Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Const WM_SYSCOMMAND = &H112

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Enum DrawTypes
    DI_MASK = &H1
    DI_IMAGE = &H2
    di_normal = &H3
    DI_COMPAT = &H4
    DI_DEFAULTSIZE = &H8
End Enum

' GetSystemMetrics() codes
Public Enum SystemMetrics
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_CMETRICS = 44
End Enum

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const HTMINBUTTON = 8
Private Const HTMAXBUTTON = 9
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_SIZE = &H5
Private Const WM_ACTIVATE = &H6
Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_GETMINMAXINFO As Long = &H24
Private MinX, MinY, MaxX, MaxY As Long

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
'Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_CALCRECT = &H400
Private Const DT_WORD_ELLIPSIS = &H40000
    
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const PS_SOLID = 0
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   itemData As Long
End Type

Private Const WM_DRAWITEM = &H2B

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)
    

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000            ' WS_BORDER Or WS_DLGFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SYSMENU = &H80000

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20 ' The frame changed: send WM_NCCALCSIZE
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200 ' Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Public Enum SysColors
    COLOR_SCROLLBAR = 0
    COLOR_BACKGROUND = 1
    COLOR_ACTIVECAPTION = 2
    COLOR_INACTIVECAPTION = 3
    COLOR_MENU = 4
    COLOR_WINDOW = 5
    COLOR_WINDOWFRAME = 6
    COLOR_MENUTEXT = 7
    COLOR_WINDOWTEXT = 8
    COLOR_CAPTIONTEXT = 9
    COLOR_ACTIVEBORDER = 10
    COLOR_INACTIVEBORDER = 11
    COLOR_APPWORKSPACE = 12
    COLOR_HIGHLIGHT = 13
    COLOR_HIGHLIGHTTEXT = 14
    COLOR_BTNFACE = 15
    COLOR_BTNSHADOW = 16
    COLOR_GRAYTEXT = 17
    COLOR_BTNTEXT = 18
    COLOR_INACTIVECAPTIONTEXT = 19
    COLOR_BTNHIGHLIGHT = 20
End Enum

Public Enum WindowCommands
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_MAX = 10
End Enum

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_ERASE = &H4

Private Const RDW_VALIDATE = &H8
Private Const RDW_NOINTERNALPAINT = &H10
Private Const RDW_NOERASE = &H20

Private Const RDW_NOCHILDREN = &H40
Private Const RDW_ALLCHILDREN = &H80

Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASENOW = &H200

Private Const RDW_FRAME = &H400
Private Const RDW_NOFRAME = &H800

Implements ISubclass
Private m_emr As EMsgResponse

Dim Ghwnd As Long
Dim ButtonBackColor As Long
Dim Textcolor As Long
Dim HilightTextcolor As Long
Dim DisabledTextColor As Long
Dim BtnHilightColor As Long
Dim BtnShadowColor As Long
Dim TranSP As Boolean
Dim DrawFocus As Boolean
Dim TitleBarHdc As Long
Dim InMenuLoop As Boolean
Dim ComingOut As Boolean
Dim Hover As Boolean
Dim TheCaption As String
Dim OldStyle As Boolean
Dim BackGroundColor As Long
Dim OwnBackColor As Boolean
Dim PicHandle As StdPicture
Dim PicYesNo As Boolean
Dim ShadeColour As Long
Dim PictType As Integer
Dim IconWidth As Long
Dim IconHeight As Long
Dim Ghdc As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Function HookAddress(ByVal lPtr As Long) As Long
   ' Work around for VB's poor AddressOf implementation:
   HookAddress = lPtr
End Function

Public Sub DrawSysButton(DrawDc As Long, lT As Long, rt As Long, tp As Long, bt As Long, styl As SysButtons, Optional inact As SysButtonsI = SysButtonsI.DFCS_ACTIVE)
Dim Frct As RECT
Frct.Left = lT
Frct.Right = rt + lT
Frct.Top = tp
Frct.Bottom = bt + tp
DrawFrameControl DrawDc, Frct, DFC_CAPTION, styl Or inact
End Sub

Public Function GetSysMetrics(met As SystemMetrics) As Long
GetSysMetrics = GetSystemMetrics(met)
End Function

Public Sub SystemMenu(frm As Form, xpos As Long, ypos As Long)
Dim lprc As RECT
Call SendMessage(frm.hwnd, WM_GETSYSMENU, 0, ByVal MAKELONG(xpos, ypos))
End Sub

Public Sub cDrawIcon(dest As Object, Source As Form, xx As Long, yy As Long, xWidth, yHeight, style As DrawTypes)
Dim TheIcon As Long
TheIcon = Source.Icon
DrawIconEx dest.hdc, xx, yy, TheIcon, xWidth, yHeight, ByVal 0&, ByVal 0&, style
End Sub

Public Sub cDrawText(obj As Object, txt As String, lT As Long, rt As Long, tp As Long, bt As Long)
Dim rctmp As RECT
Dim tmprect As RECT
'rctmp.Top = tp
'rctmp.Bottom = bt
'rctmp.Left = lT
'rctmp.Right = rt
SetRect rctmp, lT, tp, rt, bt

tmprect = rctmp

DrawText obj.hdc, txt, Len(txt), rctmp, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_CALCRECT

'tmprect.Top = tmprect.Bottom / 2 - rctmp.Bottom / 2
'tmprect.Bottom = tmprect.Bottom + tmprect.Top
'tmprect.Right = rctmp.Right
'tmprect.Left = rctmp.Left
SetRect tmprect, rctmp.Left, _
        tmprect.Bottom / 2 - rctmp.Bottom / 2, _
        rctmp.Right, _
        tmprect.Bottom + tmprect.Top

DrawText obj.hdc, txt, Len(txt), tmprect, DT_LEFT Or _
    DT_SINGLELINE Or DT_WORD_ELLIPSIS

End Sub

Public Sub WindowsCommands(frm As Form, cmd As WindowCommands)
ShowWindow frm.hwnd, cmd
End Sub

Public Sub RepaintWindow()
Dim FrmTmp As Form
For Each FrmTmp In Forms
    If FrmTmp.hwnd = Ghwnd Then
        FrmTmp.Visible = False
        FrmTmp.Visible = True
        Exit For
    End If
Next
End Sub

Public Sub RefreshButtons(Optional butt As CommandButton)
On Error Resume Next
Dim cmd As CommandButton
If butt <> False Then
'    Debug.Print "TRUE"
Dim FrmTmp As Form
Dim cmdtmp As Object
For Each FrmTmp In Forms
    If FrmTmp.hwnd = Ghwnd Then
        For Each cmdtmp In FrmTmp
        'If cmdtmp = cmd Then
        If TypeOf cmdtmp Is CommandButton Then
        cmdtmp.Refresh
        End If
        Next
        Exit Sub
    End If
Next
Else
    butt.Refresh
    Exit Sub
End If
End Sub

Function SysColour(col As SysColors)
SysColour = GetSysColor(col)
End Function

Private Sub Class_Initialize()
OwnBackColor = True
ButtonBackColor = SysColour(COLOR_BTNFACE)
Textcolor = SysColour(COLOR_BTNTEXT)
BtnHilightColor = SysColour(COLOR_BTNHIGHLIGHT)
BtnShadowColor = SysColour(COLOR_BTNSHADOW)
HilightTextcolor = QBColor(4)
DisabledTextColor = QBColor(8)
ShadeColour = QBColor(15)
HoverStyle = False
MinX = 1
MinY = 1
MaxX = Screen.Width / 15
MaxY = Screen.Height / 15

Set m_t1 = New CTimer
m_t1.Interval = 0
End Sub

Public Property Get HoverStyle() As Boolean
    HoverStyle = Hover
End Property

Public Property Let HoverStyle(ans As Boolean)
Hover = ans

If Hover = True And Addr = 0 Then
    Addr = SetWindowsHookEx(WH_MOUSE, HookAddress(AddressOf MouseProc), 0&, GetCurrentThreadId)
Else
    UnhookWindowsHookEx Addr
    Addr = 0
End If
End Property

Public Property Get ShadeColor() As Long
    ShadeColor = ShadeColour
End Property

Public Property Let ShadeColor(ans As Long)
    ShadeColour = ans
End Property

Public Property Get OwnBackColour() As Boolean
    OwnBackColour = OwnBackColor
End Property

Public Property Let OwnBackColour(ans As Boolean)
    OwnBackColor = ans
End Property

Public Property Get DrawFocusRct() As Boolean
    DrawFocusRct = DrawFocus
End Property

Public Property Let DrawFocusRct(ans As Boolean)
    DrawFocus = ans
End Property

Public Property Get OldStyleSystemButtons() As Boolean
    OldStyleSystemButtons = OldStyle
End Property

Public Property Let OldStyleSystemButtons(ans As Boolean)
    OldStyle = ans
End Property
Public Property Get TransparentButton() As Boolean
    TransparentButton = TranSP
End Property

Public Property Let TransparentButton(ans As Boolean)
    TranSP = ans
End Property

Public Property Get DisabledTextColour() As Long
    DisabledTextColour = DisabledTextColor
End Property

Public Property Let DisabledTextColour(TheColor As Long)
    DisabledTextColor = TheColor
End Property

Public Property Get BackColour() As Long
    BackColour = ButtonBackColor
End Property

Public Property Let BackColour(TheColor As Long)
    ButtonBackColor = TheColor
End Property

Public Property Get TextColour() As Long
    TextColour = Textcolor
End Property

Public Property Let TextColour(TheColor As Long)
    Textcolor = TheColor
End Property

Public Property Get BorderHilightColour() As Long
    BorderHilightColour = BtnHilightColor
End Property

Public Property Let BorderHilightColour(TheColor As Long)
    BtnHilightColor = TheColor
End Property

Public Property Get BorderShadowColour() As Long
    BorderShadowColour = BtnShadowColor
End Property

Public Property Let BorderShadowColour(TheColor As Long)
    BtnShadowColor = TheColor
End Property

Public Property Get FocusTextColour() As Long
    FocusTextColour = HilightTextcolor
End Property

Public Property Let FocusTextColour(TheColor As Long)
    HilightTextcolor = TheColor
End Property

Private Function ShowTitleBar(frm As Form, ByVal bState As Boolean)
Dim lStyle As Long
Dim tR As RECT

    ' Get the window's position:
    GetWindowRect Ghwnd, tR

    ' Modify whether title bar will be visible:
    lStyle = GetWindowLong(Ghwnd, GWL_STYLE)
    If (bState) Then
    If frm.ControlBox Then
        lStyle = lStyle Or WS_SYSMENU
    End If
    If frm.MaxButton Then
        lStyle = lStyle Or WS_MAXIMIZEBOX
    End If
    If frm.MinButton Then
        lStyle = lStyle Or WS_MINIMIZEBOX
    End If
    If frm.Caption <> "" Then
        lStyle = lStyle Or WS_CAPTION
    End If
    Else
    lStyle = lStyle And Not WS_MAXIMIZEBOX
    lStyle = lStyle And Not WS_MINIMIZEBOX
    lStyle = lStyle And Not WS_CAPTION
End If
SetWindowLong Ghwnd, GWL_STYLE, lStyle

' Ensure the style takes and make the window the
' same size, regardless that the title bar etc
' is now a different size:
SetWindowPos Ghwnd, 0, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
frm.Refresh
End Function

Private Function FindControl(frm As Form, thwnd As Long) As String
PicYesNo = False
Dim ctrl As Control
For Each ctrl In frm
        If ctrl.hwnd = thwnd Then
            FindControl = ctrl.Tag
            TheCaption = ctrl.Caption
            BackGroundColor = ctrl.BackColor
            If ctrl.Picture > 0 And UCase(ctrl.Tag) <> "NORMAL" Then
                'IconWidth = Int(ctrl.Picture.Width / 26)
                'IconHeight = Int(ctrl.Picture.Height / 26)
                IconWidth = frm.ScaleX(ctrl.Picture.Width, vbHimetric, vbPixels)
                IconHeight = frm.ScaleY(ctrl.Picture.Height, vbHimetric, vbPixels)
                Set PicHandle = ctrl.Picture
                PicYesNo = True
                PictType = ctrl.Picture.Type
                'Set m_pic = ctrl.Picture
                'Module2.pbGetBitmapIntoDC
            End If
            Exit For
        End If
Next
End Function

Sub MakeCustomButtons(frm As Object, Optional titlebar As Boolean = True)
Ghwnd = frm.hwnd
Ghdc = frm.hdc

If titlebar = False Then
    ShowTitleBar frm, False
End If

AttachMessage Me, Ghwnd, WM_DRAWITEM
AttachMessage Me, Ghwnd, WM_GETMINMAXINFO
AttachMessage Me, Ghwnd, WM_SIZE
AttachMessage Me, Ghwnd, WM_ACTIVATE
AttachMessage Me, Ghwnd, WM_CUSTOM
AttachMessage Me, Ghwnd, WM_SYSCOLORCHANGE

If HoverStyle And Addr = 0 Then
    Addr = SetWindowsHookEx(WH_MOUSE, HookAddress(AddressOf MouseProc), 0&, GetCurrentThreadId)
End If
End Sub

Sub UnMakeCustomButtons(frm As Object)
Ghwnd = frm.hwnd
DetachMessage Me, Ghwnd, WM_DRAWITEM
DetachMessage Me, Ghwnd, WM_GETMINMAXINFO
DetachMessage Me, Ghwnd, WM_SIZE
DetachMessage Me, Ghwnd, WM_ACTIVATE
DetachMessage Me, Ghwnd, WM_CUSTOM
DetachMessage Me, Ghwnd, WM_SYSCOLORCHANGE

If HoverStyle And Addr <> 0 Then
    UnhookWindowsHookEx Addr
End If
End Sub

Sub MakeTitleBar(obj As Object)
AttachMessage Me, obj.hwnd, WM_LBUTTONDOWN
AttachMessage Me, Ghwnd, WM_ENTERMENULOOP

obj.Height = GetSysMetrics(SM_CYCAPTION) - 1
TitleBarHdc = obj.hdc
End Sub

Sub UnMakeTitleBar(obj As Object)
DetachMessage Me, obj.hwnd, WM_LBUTTONDOWN
DetachMessage Me, Ghwnd, WM_ENTERMENULOOP
TitleBarHdc = 0
End Sub

Sub SetMinMaxInfo(MinXc As Long, MinYc As Long, MaxXc As Long, MaxYc As Long)
    MinX = MinXc
    MinY = MinYc
    MaxX = MaxXc
    MaxY = MaxYc
End Sub

Private Sub DrawSystemButtons(hwnd As Long)
Dim TmpHeight As Long
Dim TmpWnd As RECT
Dim TmpFrame As Long

GetWindowRect hwnd, TmpWnd

TmpHeight = GetSysMetrics(SM_CYCAPTION)
TmpFrame = GetSysMetrics(SM_CXDLGFRAME) + 1

DrawSysButton TitleBarHdc, _
    TmpWnd.Right - TmpWnd.Left - TmpHeight - 1 - TmpFrame _
    , TmpHeight - 2 - 1, 2, _
    TmpHeight - 4 - 1, DFCS_CAPTIONCLOSE

Dim FrmTmp As Form
Dim resizetype As Long
For Each FrmTmp In Forms
If FrmTmp.hwnd = hwnd Then
    resizetype = FrmTmp.WindowState
    Exit For
End If
Next

If resizetype = 0 Then
DrawSysButton TitleBarHdc, _
    TmpWnd.Right - TmpWnd.Left - (TmpHeight * 2) - TmpFrame _
    , TmpHeight - 2 - 1, 2, _
    TmpHeight - 4 - 1, DFCS_CAPTIONMAX
Else
DrawSysButton TitleBarHdc, _
    TmpWnd.Right - TmpWnd.Left - (TmpHeight * 2) - TmpFrame _
    , TmpHeight - 2 - 1, 2, _
    TmpHeight - 4 - 1, DFCS_CAPTIONRESTORE
End If

DrawSysButton TitleBarHdc, _
    TmpWnd.Right - TmpWnd.Left - (TmpHeight * 3) + 2 - TmpFrame + 1 _
    , TmpHeight - 2 - 1, 2, _
    TmpHeight - 4 - 1, DFCS_CAPTIONMIN

End Sub

Private Sub DrawTransparent(lHDC As Long, lHwnd As Long, rItem As RECT)
Dim cctmp As Object
Dim ccfrm As Form
For Each ccfrm In Forms
'    If ccfrm.hwnd = Ghwnd Then
        For Each cctmp In ccfrm
            If cctmp.hwnd = lHwnd Then
                BitBlt lHDC, 0, 0, rItem.Right, _
                    rItem.Bottom, Ghdc, _
                    cctmp.Left, cctmp.Top, vbSrcCopy
                Exit Sub
            End If
        Next
'        Exit For
'    End If
Next
End Sub

Sub ClickTitleBar()
Call ReleaseCapture
SendMessage Ghwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

If iMsg = WM_SYSCOLORCHANGE Then
Dim frmtm As Form
For Each frmtm In Forms
If frmtm.hwnd = hwnd Then
    If GetActiveWindow = hwnd Then
        RaiseEvent Resizing(Ghwnd, True, frmtm.WindowState, True)
    Else
        RaiseEvent Resizing(Ghwnd, False, frmtm.WindowState, True)
    End If
    CallOldWindowProc hwnd, iMsg, wParam, lParam
    Exit Function
End If
Next
End If

If iMsg = WM_CUSTOM Then
    If lParam > 0 Then
        m_t1.Interval = 50
    ElseIf lParam = 0 Then
        m_t1.Interval = 0
    End If
Exit Function
End If

If iMsg = WM_ACTIVATE And hwnd = Ghwnd Then
Debug.Print "Activate " & wParam & " " & lParam
    If CBool(wParam) = True Then
        RaiseEvent Resizing(lParam, True, 0, False)
    ElseIf CBool(wParam) = False Then
        RaiseEvent Resizing(lParam, False, 0, False)
    End If
    CallOldWindowProc hwnd, iMsg, wParam, lParam
    Exit Function
End If

If iMsg = WM_SIZE And hwnd = Ghwnd Then
    RaiseEvent Resizing(Ghwnd, True, wParam, False)
    Exit Function
End If

If iMsg = WM_GETMINMAXINFO And hwnd = Ghwnd Then
    Dim MMI As MINMAXINFO
    CopyMemory MMI, ByVal lParam, LenB(MMI)
    With MMI
        .ptMinTrackSize.x = MinX
        .ptMinTrackSize.y = MinY
        .ptMaxTrackSize.x = MaxX
        .ptMaxTrackSize.y = MaxY
    End With
    CopyMemory ByVal lParam, MMI, LenB(MMI)
    Exit Function
End If

If iMsg = WM_MEASUREITEM Then
Dim mM As MEASUREITEMSTRUCT
CopyMemory mM, lParam, Len(mM)
mM.itemHeight = 50
mM.itemWidth = 100
CopyMemory lParam, mM, Len(mM)
Exit Function
End If

If iMsg = WM_DRAWITEM Then

    Dim tDis As DRAWITEMSTRUCT
    Dim hBr As Long
    Dim tJunk As POINTAPI
    Dim Frct As RECT
    Dim ButPict As Boolean
    Dim DrawType As Long
    
    CopyMemory tDis, ByVal lParam, Len(tDis)

'If control not a command button then exit function
If tDis.CtlType = 1 Then
CallOldWindowProc hwnd, iMsg, wParam, lParam
Exit Function
End If

    Dim rtmp As Long
    Dim Classtmp As String
    Dim tmphwnd As Long
    Classtmp = Space(255)
    rtmp = GetClassName(hwnd, Classtmp, 255)
    
    If Classtmp <> "ThunderFormDC" Then
        tmphwnd = GetParent(hwnd)
        If tmphwnd = 0 Then tmphwnd = hwnd
    Else
        tmphwnd = hwnd
    End If
    
    Dim TmpControl As String
    Dim tmpform As Form
    For Each tmpform In Forms
        If tmpform.hwnd = tmphwnd Then
            TmpControl = FindControl(tmpform, tDis.hwndItem)
            Exit For
        End If
    Next

    If UCase(TmpControl) = "NORMAL" Then
        CallOldWindowProc hwnd, iMsg, wParam, lParam
        If m_lHdc > 0 Then
            Module2.pClearUp
        End If
        Exit Function
    End If

    'If the command button has a tag="SYSbutton" then
    'fill colour is the system colour else use the
    'custom colour.
    If TmpControl <> "SYSbutton" Then
        If BackGroundColor < 0 Or Not OwnBackColor Then
            hBr = CreateSolidBrush(ButtonBackColor)
        Else
            hBr = CreateSolidBrush(BackGroundColor)
        End If
    Else
        hBr = CreateSolidBrush(SysColour(COLOR_BTNFACE))
    End If
    
    'Fill only fill the button if it have a tag of
    '"SYSbutton" or the TranSP is set to TRUE(If
    'the button is transparent then it would be a
    'waste of time to fill it in).
    If TranSP = False Or TmpControl = "SYSbutton" Then
        FillRect tDis.hdc, tDis.rcItem, hBr
    End If
    DeleteObject hBr
   
    If TmpControl <> "SYSbutton" And TranSP = True Then
        DrawTransparent tDis.hdc, tDis.hwndItem, tDis.rcItem
    End If
    
    'Get the buttons caption.
    Dim sText As String
    Dim sTexttmp As String
    sText = Space(255)
    sTexttmp = GetWindowText(tDis.hwndItem, sText, 255)
    sText = Trim$(sText)
    sText = Mid(sText, 1, Len(sText) - 1)
    
    Dim tmprect As RECT
    tmprect = tDis.rcItem
   
    'Make the text background transparent.
    SetBkMode tDis.hdc, 1

    If tDis.itemState = 0 And TmpControl <> "SYSbutton" Then
        SetTextColor tDis.hdc, Textcolor
    ElseIf tDis.itemState = 16 And TmpControl <> "SYSbutton" Then
        SetTextColor tDis.hdc, HilightTextcolor
   
    'New focus rect
    If DrawFocus Then
        Frct = tDis.rcItem
        InflateRect Frct, -4, -4
        DrawFocusRect tDis.hdc, Frct
    End If

    ElseIf tDis.itemState = 4 Then
        SetTextColor tDis.hdc, DisabledTextColor
    ElseIf tDis.itemState = 17 Then
        SetTextColor tDis.hdc, HilightTextcolor
        If TmpControl = "SYSbutton" Then
            SetTextColor tDis.hdc, SysColour(COLOR_BTNTEXT)
        End If

    'Draws a focus rectangle with the right settings.
    If TmpControl <> "SYSbutton" And DrawFocus Then
        Frct = tDis.rcItem
        InflateRect Frct, -4, -4
        DrawFocusRect tDis.hdc, Frct
    End If
    End If
      
    If TmpControl <> "SYSbutton" Then
    DrawText tDis.hdc, sText, -1, tmprect, DT_WORDBREAK Or DT_CALCRECT

    tmprect.Top = tDis.rcItem.Bottom / 2 - _
                tmprect.Bottom / 2
    tmprect.Bottom = tmprect.Bottom + tmprect.Top

    tmprect.Left = tDis.rcItem.Right / 2 - _
                tmprect.Right / 2
    tmprect.Right = tmprect.Right + tmprect.Left

'Starts here
If m_lHdc > 0 Or PicYesNo = True Then

Dim tmpbottom As Long
tmpbottom = tmprect.Bottom - tmprect.Top
tmprect.Top = (tDis.rcItem.Bottom - tDis.rcItem.Top) - tmpbottom - 2
tmprect.Bottom = tmprect.Top + tmpbottom

Dim xtmp, ytmp As Long

Dim tB2 As BITMAP
GetObjectAPI PicHandle.handle, LenB(tB2), tB2

If PictType = 1 Then
    DrawType = DST_BITMAP
    xtmp = Int((tDis.rcItem.Right / 2) - (tB2.bmWidth / 2))
    ytmp = Int((tmprect.Top / 2) - (tB2.bmHeight / 2))
ElseIf PictType = 3 Then
    DrawType = DST_ICON
    xtmp = Int((tDis.rcItem.Right / 2) - (IconWidth / 2))
    ytmp = Int((tmprect.Top / 2) - (IconHeight / 2))
End If

If tDis.itemState = 17 Then
    xtmp = xtmp + 1
    ytmp = ytmp + 1
End If

If tDis.itemState = 4 Then
Dim Gh As Long
Gh = CreateSolidBrush(ShadeColour)
DrawState tDis.hdc, 0, 0, PicHandle, 0, _
    xtmp, ytmp, _
    tB2.bmWidth, tB2.bmHeight, _
    DrawType 'Or DSS_MONO Or DSS_UNION
DrawState tDis.hdc, Gh, 0, PicHandle, 0, _
    xtmp, ytmp, _
    tB2.bmWidth, tB2.bmHeight, _
    DrawType Or DSS_MONO Or DSS_UNION
'frms.Refresh
DeleteObject Gh

Else
DrawState tDis.hdc, 0, 0, PicHandle, 0, _
    xtmp, ytmp, _
    tB2.bmWidth, tB2.bmHeight, _
    DrawType
End If

End If

    End If

    If tDis.itemState = 17 Then
        tmprect.Left = tmprect.Left + 1
        tmprect.Top = tmprect.Top + 1
        tmprect.Right = tmprect.Right + 1
        tmprect.Bottom = tmprect.Bottom + 1
    End If

If TmpControl <> "SYSbutton" Then
    DrawText tDis.hdc, sText, -1, tmprect, DT_CENTER Or DT_WORDBREAK Or DT_WORD_ELLIPSIS
End If
    
If TmpControl = "SYSbutton" Then
Dim StyleEx
StyleEx = &H0
    If tDis.itemState = 17 Then
    StyleEx = &H200
    ElseIf tDis.itemState = 4 Then
    StyleEx = &H100
    End If
        If TheCaption = "r" Then
        DrawFrameControl tDis.hdc, tDis.rcItem, DFC_CAPTION, SysButtons.DFCS_CAPTIONCLOSE Or StyleEx
        ElseIf TheCaption = "1" Then
        DrawFrameControl tDis.hdc, tDis.rcItem, DFC_CAPTION, SysButtons.DFCS_CAPTIONMAX Or StyleEx
        ElseIf TheCaption = "0" Then
        DrawFrameControl tDis.hdc, tDis.rcItem, DFC_CAPTION, SysButtons.DFCS_CAPTIONMIN Or StyleEx
        ElseIf TheCaption = "2" Then
        DrawFrameControl tDis.hdc, tDis.rcItem, DFC_CAPTION, SysButtons.DFCS_CAPTIONRESTORE Or StyleEx
        End If
End If
    
    Dim DarkGrey, LightGrey As Long
    
    If tDis.itemState = 16 _
        Or tDis.itemState = 0 _
        Or tDis.itemState = 4 Then
    If TmpControl <> "SYSbutton" Then
        DarkGrey = CreatePen(PS_SOLID, 1, BtnShadowColor)
        LightGrey = CreatePen(PS_SOLID, 1, BtnHilightColor)
    Else
        DarkGrey = CreatePen(PS_SOLID, 1, SysColour(COLOR_BTNSHADOW))
        LightGrey = CreatePen(PS_SOLID, 1, SysColour(COLOR_BTNHIGHLIGHT))
    End If
    ElseIf tDis.itemState = 17 Then
    'Button Down
    If TmpControl <> "SYSbutton" Then
        DarkGrey = CreatePen(PS_SOLID, 1, BtnHilightColor)
        LightGrey = CreatePen(PS_SOLID, 1, BtnShadowColor)
    Else
        DarkGrey = CreatePen(PS_SOLID, 1, SysColour(COLOR_BTNHIGHLIGHT))
        LightGrey = CreatePen(PS_SOLID, 1, SysColour(COLOR_BTNSHADOW))
    End If
    End If
      
    If HoverStyle = True Then
    If MouseHwnd = tDis.hwndItem Or TmpControl = "SYSbutton" Then
        If TmpControl = "SYSbutton" And OldStyle Then GoTo enn
        SelectObject tDis.hdc, LightGrey
        LineTo tDis.hdc, tDis.rcItem.Right - 1, 0
        SelectObject tDis.hdc, DarkGrey
        LineTo tDis.hdc, tDis.rcItem.Right - 1, tDis.rcItem.Bottom - 1
        LineTo tDis.hdc, 0, tDis.rcItem.Bottom - 1
        SelectObject tDis.hdc, LightGrey
        LineTo tDis.hdc, 0, 0
    End If
    ElseIf HoverStyle = False Then
        If TmpControl = "SYSbutton" And OldStyle Then GoTo enn
        SelectObject tDis.hdc, LightGrey
        LineTo tDis.hdc, tDis.rcItem.Right - 1, 0
        SelectObject tDis.hdc, DarkGrey
        LineTo tDis.hdc, tDis.rcItem.Right - 1, tDis.rcItem.Bottom - 1
        LineTo tDis.hdc, 0, tDis.rcItem.Bottom - 1
        SelectObject tDis.hdc, LightGrey
        LineTo tDis.hdc, 0, 0
    End If
    
enn:
    If TmpControl = "SYSbutton" And Not OldStyle Then
        Dim tmpPen As Long
        Dim TmpPnt As POINTAPI
        tmpPen = CreatePen(PS_SOLID, 1, SysColour(COLOR_BTNFACE))
        MoveToEx tDis.hdc, 1, 1, TmpPnt
        SelectObject tDis.hdc, tmpPen
        LineTo tDis.hdc, tDis.rcItem.Right - 2, 1
        LineTo tDis.hdc, tDis.rcItem.Right - 2, tDis.rcItem.Bottom - 2
        LineTo tDis.hdc, 1, tDis.rcItem.Bottom - 2
        LineTo tDis.hdc, 1, 1
        DeleteObject tmpPen
    End If
        
    ReleaseDC tDis.hwndItem, tDis.hdc
    DeleteObject LightGrey
    DeleteObject DarkGrey
End If
End Function

Private Sub m_t1_ThatTime()
Dim WndTmp As RECT
Dim MP As POINTAPI
GetWindowRect Ghwnd, WndTmp
GetCursorPos MP

If MP.x < WndTmp.Left Or _
    MP.x > WndTmp.Right Or _
    MP.y < WndTmp.Top Or _
    MP.y > WndTmp.Bottom Then
RedrawWindow MouseHwnd, WndTmp, RDW_UPDATENOW, 1
MouseHwnd = 0
m_t1.Interval = 0
End If
End Sub

Private Function LOWORD(dwValue As Long) As Long
    CopyMemory LOWORD, dwValue, 2
End Function

Private Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHigh))
End Function
