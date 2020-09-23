Attribute VB_Name = "Module1"
Option Explicit
Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long

Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Const RDW_INVALIDATE = &H1
Public Const RDW_INTERNALPAINT = &H2
Public Const RDW_ERASE = &H4

Public Const RDW_VALIDATE = &H8
Public Const RDW_NOINTERNALPAINT = &H10
Public Const RDW_NOERASE = &H20

Public Const RDW_NOCHILDREN = &H40
Public Const RDW_ALLCHILDREN = &H80

Public Const RDW_UPDATENOW = &H100
Public Const RDW_ERASENOW = &H200

Public Const RDW_FRAME = &H400
Public Const RDW_NOFRAME = &H800


Const WM_MOUSEMOVE = &H200

Type POINTAPI
        x As Long
        y As Long
End Type

Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hwnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type

Public Const WH_MOUSE = 7

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowsHook Lib "user32" Alias "SetWindowsHookA" (ByVal nFilterType As Long, ByVal pfnFilterProc As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHook Lib "user32" (ByVal nCode As Long, ByVal pfnFilterProc As Long) As Boolean
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long



Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Public Enum DrawTyps
    dDI_MASK = &H1
    dDI_IMAGE = &H2
    ddi_normal = &H3
    dDI_COMPAT = &H4
    dDI_DEFAULTSIZE = &H8
End Enum

Private Type PictDesc
   cbSizeofStruct As Long
   picType As Long
   hImage As Long
   xExt As Long
   yExt As Long
End Type

Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
        (lpPictDesc As PictDesc, riid As Guid, _
        ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function BitBltt Lib "gdi32" Alias "BitBlt" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)

'Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
'Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public PictureTmp As Long
Global MouseHwnd As Long
Global Addr As Long
Global RefreshWindow As RECT
'Public Stage As Long

Function NewDC(hdcScreen As Long, HorRes As Long, VerRes As Long) As Long
    Dim hdcCompatible As Long
    Dim hbmScreen As Long
    hdcCompatible = CreateCompatibleDC(hdcScreen)                   'Create the DC
    hbmScreen = CreateCompatibleBitmap(hdcScreen, HorRes, VerRes)   'Temporary bitmap
    If SelectObject(hdcCompatible, hbmScreen) = vbNull Then         'If the function fails
        NewDC = vbNull                                              ' return null
    Else                                                            'If it succeeds
        NewDC = hdcCompatible                                       ' return the DC
    End If
End Function

Public Function BitmapToPicture(ByVal hBmp As Long) As IPicture

   If (hBmp = 0) Then Exit Function
   Dim oNewPic As Picture, tPicConv As PictDesc, IGuid As Guid
   ' Fill PictDesc structure with necessary parts:
   With tPicConv
      .cbSizeofStruct = Len(tPicConv)
      .picType = vbPicTypeBitmap
      .hImage = hBmp
   End With
   ' Fill in IDispatch Interface ID
   With IGuid
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   ' Create a picture object:
   OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
   ' Return it:
   Set BitmapToPicture = oNewPic

End Function
'-- End --'

Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next

Dim MouseTmp As MOUSEHOOKSTRUCT
CopyMemory MouseTmp, ByVal lParam, Len(MouseTmp)

If WM_MOUSEMOVE Then
    Dim pHwnd As Long
    pHwnd = WindowFromPoint(MouseTmp.pt.x, MouseTmp.pt.y)
'Debug.Print "HWND " & pHwnd

If pHwnd <> MouseHwnd Then
Dim OldHwnd As Long
Dim NewR, OldR As Boolean
    OldHwnd = MouseHwnd
    MouseHwnd = pHwnd

Dim ctmp As Form
For Each ctmp In Forms
    If ctmp.hwnd = pHwnd Then NewR = True: Exit For
    If ctmp.hwnd = OldHwnd Then OldR = True: Exit For
Next

Dim Wrect As RECT

If NewR = False And pHwnd > 0 Then
    GetWindowRect pHwnd, Wrect
    RefreshWindow = Wrect
    RedrawWindow pHwnd, Wrect, RDW_UPDATENOW, 1
    SendMessage ctmp.hwnd, &H400, 1, 1
End If
    
If OldR = False And OldHwnd > 0 Then
    GetWindowRect OldHwnd, Wrect
    RedrawWindow OldHwnd, Wrect, RDW_UPDATENOW, 1
End If

End If
End If

MouseProc = CallNextHookEx(WH_MOUSE, nCode, wParam, lParam)
End Function




