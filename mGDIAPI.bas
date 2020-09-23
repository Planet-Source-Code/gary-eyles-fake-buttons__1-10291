Attribute VB_Name = "mGDIAPI"
Option Explicit

' General:
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
    Public Const GWW_HINSTANCE = (-6)
    
' GDI object functions:
'Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'    Public Const BITSPIXEL = 12
' System metrics:
'Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'    Public Const SM_CXICON = 11
'    Public Const SM_CYICON = 12
'    Public Const SM_CXFRAME = 32
'    Public Const SM_CYCAPTION = 4
'    Public Const SM_CYFRAME = 33
'    Public Const SM_CYBORDER = 6
'    Public Const SM_CXBORDER = 5

' Region paint and fill functions:
'Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
'Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'    Public Const FLOODFILLBORDER = 0
'    Public Const FLOODFILLSURFACE = 1

' Pen functions:
'Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'    Public Const PS_DASH = 1
'    Public Const PS_DASHDOT = 3
'    Public Const PS_DASHDOTDOT = 4
'    Public Const PS_DOT = 2
'    Public Const PS_SOLID = 0
'    Public Const PS_NULL = 5
'
' Brush functions:
'Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

' Line functions:
'Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Type POINTAPI
'        x As Long
'        y As Long
'End Type
'Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function FillPath Lib "gdi32" (ByVal hdc As Long) As Long

' Colour functions:
'Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
'Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
'    Public Const OPAQUE = 2
'    Public Const TRANSPARENT = 1
'Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'    Public Const COLOR_ACTIVEBORDER = 10
'    Public Const COLOR_ACTIVECAPTION = 2
'    Public Const COLOR_ADJ_MAX = 100
'    Public Const COLOR_ADJ_MIN = -100
'    Public Const COLOR_APPWORKSPACE = 12
'    Public Const COLOR_BACKGROUND = 1
'    Public Const COLOR_BTNFACE = 15
'    Public Const COLOR_BTNHIGHLIGHT = 20
'    Public Const COLOR_BTNSHADOW = 16
'    Public Const COLOR_BTNTEXT = 18
'    Public Const COLOR_CAPTIONTEXT = 9
'    Public Const COLOR_GRAYTEXT = 17
'    Public Const COLOR_HIGHLIGHT = 13
'    Public Const COLOR_HIGHLIGHTTEXT = 14
'    Public Const COLOR_INACTIVEBORDER = 11
'    Public Const COLOR_INACTIVECAPTION = 3
'    Public Const COLOR_INACTIVECAPTIONTEXT = 19
'    Public Const COLOR_MENU = 4
'    Public Const COLOR_MENUTEXT = 7
    'Public Const COLOR_SCROLLBAR = 0
'    Public Const COLOR_WINDOW = 5
'    Public Const COLOR_WINDOWFRAME = 6
'    Public Const COLOR_WINDOWTEXT = 8
'    Public Const COLORONCOLOR = 3

' Shell Extract icon functions:
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

' GDI icon functions:
'Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' Blitting functions
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCAND = &H8800C6
    Public Const SRCCOPY = &HCC0020
    Public Const SRCERASE = &H440328
    Public Const SRCINVERT = &H660046
    Public Const SRCPAINT = &HEE0086
    Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062
'Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
'Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
'Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Type BITMAP
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type
Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long

' Text functions:
'Type RECT
'    Left As Long
'    TOp As Long
'    Right As Long
'    Bottom As Long
'End Type
Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
'Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'    Public Const DT_BOTTOM = &H8
'    Public Const DT_CENTER = &H1
'    Public Const DT_LEFT = &H0
'    Public Const DT_CALCRECT = &H400
'    Public Const DT_WORDBREAK = &H10
'    Public Const DT_VCENTER = &H4
'    Public Const DT_TOP = &H0
'    Public Const DT_TABSTOP = &H80
'    Public Const DT_SINGLELINE = &H20
 '   Public Const DT_RIGHT = &H2
 ''   Public Const DT_NOCLIP = &H100
 '   Public Const DT_INTERNAL = &H1000
 '   Public Const DT_EXTERNALLEADING = &H200
 '   Public Const DT_EXPANDTABS = &H40
 '   Public Const DT_CHARSTREAM = 4
'Declare Function ExtTextOutRect Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
'Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
'    Public Const ETO_CLIPPED = 4
'    Public Const ETO_GRAYED = 1
'    Public Const ETO_OPAQUE = 2
'    Public Const TA_BASELINE = 24
'    Public Const TA_BOTTOM = 8
'    Public Const TA_CENTER = 6
'    Public Const TA_LEFT = 0
'    Public Const TA_NOUPDATECP = 0
'    Public Const TA_UPDATECP = 1
'    Public Const TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
'    Public Const TA_RIGHT = 2
'    Public Const TA_TOP = 0

'Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
'Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

'Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'    Public Const SW_SHOWNOACTIVATE = 4

' Scrolling and region functions:
Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal hSavedDC As Long) As Long

Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Const CF_BITMAP = 2
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADTRANSPARENT = &H20
Private Const IMAGE_BITMAP = 0

'Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

