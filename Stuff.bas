Attribute VB_Name = "Module2"
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Declare Function DrawState Lib "user32" Alias "DrawStateA" _
    (ByVal hdc As Long, _
    ByVal hBrush As Long, _
    ByVal lpDrawStateProc As Long, _
    ByVal lParam As Long, _
    ByVal wParam As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal fuFlags As Long) As Long

Public Const DST_COMPLEX = &H0&
Public Const DST_TEXT = &H1&
Public Const DST_PREFIXTEXT = &H2&
Public Const DST_ICON = &H3&
Public Const DST_BITMAP = &H4&

' /* State type */
Public Const DSS_NORMAL = &H0&
Public Const DSS_UNION = &H10& ' Dither
Public Const DSS_DISABLED = &H20&
Public Const DSS_MONO = &H80& ' Draw in colour of brush specified in hBrush
Public Const DSS_RIGHT = &H8000&

Public Enum ChangeCol
BackColour = 0
ForeColour = 1
FocusForeColour = 2
DisabledTextColour = 3
BorderShadowColour = 4
BorderHilightColour = 5
DitheredColour = 6
End Enum

'Public Type RECT
'       Left As Long
'       Top As Long
'       Right As Long
'       Bottom As Long
'End Type

'Public Declare Function BitBlt Lib "gdi32" _
         (ByVal hDCDest As Long, ByVal XDest As Long, _
          ByVal YDest As Long, ByVal nWidth As Long, _
          ByVal nHeight As Long, ByVal hDCSrc As Long, _
          ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
          

Public Declare Function CreateBitmap Lib "gdi32" _
         (ByVal nWidth As Long, _
          ByVal nHeight As Long, _
          ByVal nPlanes As Long, _
          ByVal nBitCount As Long, _
          lpBits As Any) As Long

Public Declare Function SetBkColor Lib "gdi32" _
           (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Declare Function SelectObject Lib "gdi32" _
          (ByVal hdc As Long, ByVal hObject As Long) As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" _
          (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Public Declare Function CreateCompatibleDC Lib "gdi32" _
       (ByVal hdc As Long) As Long

Public Declare Function DeleteDC Lib "gdi32" _
       (ByVal hdc As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" _
       (ByVal hObject As Long) As Long

Global m_lHdc As Long
Global m_lBitmapW As Long
Global m_lBitmapH As Long
Global m_lHBmpOld As Long
Global m_lhPalOld As Long
Global m_pic As StdPicture

Function pbGetBitmapIntoDC() As Boolean
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long

    ' Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    ReleaseDC lHwnd, lHDC
    If (m_lHdc <> 0) Then
        ' Get size of bitmap:
        GetObjectAPI m_pic.handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        
        ' Select bitmap into DC:
        m_lHBmpOld = SelectObject(m_lHdc, m_pic.handle)
        If (m_lHBmpOld <> 0) Then
            ' Select the palette into the DC:
            m_lhPalOld = SelectObject(m_lHdc, m_pic.hPal)
            pbGetBitmapIntoDC = True
            If (m_sFileName = "") Then
               m_sFileName = "PICTURE"
            End If
    '    Else
    '        pClearUp
    '        pErr 2, "Unable to select bitmap into DC"
        End If
    'Else
    '    pErr 1, "Unable to create compatible DC"
    End If
End Function

Sub pClearUp()
    ' Clear reference to the filename:
    m_sFileName = ""
    ' If we have a DC, then clear up:
    If (m_lHdc <> 0) Then
        ' Select the bitmap out of DC:
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' The original bitmap does not have to deleted because it is owned by m_pic
        End If
        ' Select the palette out of the DC:
        If (m_lhPalOld <> 0) Then
            SelectObject m_lHdc, m_lhPalOld
            ' The original palette does not have to deleted because it is owned by m_pic
        End If
        ' Remove the DC:
        DeleteObject m_lHdc
    End If
m_lHdc = 0
End Sub

Sub TransBlt(OutDstDC As Long, _
       DstDC As Long, SrcDC As Long, SrcRect As RECT, _
       DstX As Integer, DstY As Integer, TransColor As Long)
       '     DstDC- Device context into which image must be drawn transparently
       '     OutDstDC- Device context into image is actually drawn, even though
       '     it is made transparent in terms of DstDC
       '     Src- Device context of source to be made transparent in color TransColor
       '     SrcRect- Rectangular region within SrcDC to be made transparent in terms of
       '     DstDC, and drawn to OutDstDC
       'DstX, DstY - Coordinates in OutDstDC (and DstDC) where the transparent bitmap m
       '     ust go
       '     In most cases, OutDstDC and DstDC will be the same
       Dim nRet As Long, W As Integer, h As Integer
       Dim MonoMaskDC As Long, hMonoMask As Long
       Dim MonoInvDC As Long, hMonoInv As Long
       Dim ResultDstDC As Long, hResultDst As Long
       Dim ResultSrcDC As Long, hResultSrc As Long
       Dim hPrevMask As Long, hPrevInv As Long
       Dim hPrevSrc As Long, hPrevDst As Long
       W = SrcRect.Right - SrcRect.Left + 1
       h = SrcRect.Bottom - SrcRect.Top + 1
       '     create monochrome mask and inverse masks
       MonoMaskDC = CreateCompatibleDC(DstDC)
       MonoInvDC = CreateCompatibleDC(DstDC)
       hMonoMask = CreateBitmap(W, h, 1, 1, ByVal 0&)
       hMonoInv = CreateBitmap(W, h, 1, 1, ByVal 0&)
       hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
       hPrevInv = SelectObject(MonoInvDC, hMonoInv)
       '     create keeper DCs and bitmaps
       ResultDstDC = CreateCompatibleDC(DstDC)
       ResultSrcDC = CreateCompatibleDC(DstDC)
       hResultDst = CreateCompatibleBitmap(DstDC, W, h)
       hResultSrc = CreateCompatibleBitmap(DstDC, W, h)
       hPrevDst = SelectObject(ResultDstDC, hResultDst)
       hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
       '     copy src to monochrome mask
       Dim OldBC As Long
       OldBC = SetBkColor(SrcDC, TransColor)
       nRet = BitBlt(MonoMaskDC, 0, 0, W, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
       TransColor = SetBkColor(SrcDC, OldBC)
       '     create inverse of mask
       nRet = BitBlt(MonoInvDC, 0, 0, W, h, MonoMaskDC, 0, 0, vbNotSrcCopy)
       '     get background
       nRet = BitBlt(ResultDstDC, 0, 0, W, h, DstDC, DstX, DstY, vbSrcCopy)
       '     AND with Monochrome mask
       nRet = BitBlt(ResultDstDC, 0, 0, W, h, MonoMaskDC, 0, 0, vbSrcAnd)
       '     get overlapper
       nRet = BitBlt(ResultSrcDC, 0, 0, W, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
       '     AND with inverse monochrome mask
       nRet = BitBlt(ResultSrcDC, 0, 0, W, h, MonoInvDC, 0, 0, vbSrcAnd)
       '     XOR these two
       nRet = BitBlt(ResultDstDC, 0, 0, W, h, ResultSrcDC, 0, 0, vbSrcInvert)
       '     output results
       nRet = BitBlt(OutDstDC, DstX, DstY, W, h, ResultDstDC, 0, 0, vbSrcCopy)
       '     clean up
       hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
       DeleteObject hMonoMask
       hMonoInv = SelectObject(MonoInvDC, hPrevInv)
       DeleteObject hMonoInv
       hResultDst = SelectObject(ResultDstDC, hPrevDst)
       DeleteObject hResultDst
       hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
       DeleteObject hResultSrc
       DeleteDC MonoMaskDC
       DeleteDC MonoInvDC
       DeleteDC ResultDstDC
       DeleteDC ResultSrcDC
End Sub

Sub DrawDithered(hwnd As Long, hdc As Long, pict As Long, xtmp, ytmp)
On Error Resume Next

Dim frms As Form
For Each frms In Forms
    If frms.hwnd = hwnd Then
        Dim objs As Object
        For Each objs In frms
            If objs.Picture = pict Then
            Exit For
            End If
        Next
    Exit For
    End If
Next

Dim Gh As Long
Gh = CreateSolidBrush(QBColor(2))
DrawState hdc, 0, 0, objs.Picture, 0, _
    xtmp, ytmp, _
    objs.Picture.Width, objs.Picture.Height, _
    DST_BITMAP 'Or DSS_MONO Or DSS_UNION
DrawState hdc, Gh, 0, objs.Picture, 0, _
    xtmp, ytmp, _
    objs.Picture.Width, objs.Picture.Height, _
    DST_BITMAP Or DSS_MONO Or DSS_UNION
'frms.Refresh
DeleteObject Gh
End Sub

Public Function CVI(s As String) As Integer
   Dim i As Integer
   
   If Len(s) <> 2 Then
      Err.Raise 1000, "CVI", "Invalid string argument"
   ElseIf Len(s) = 2 Then
      CopyMemory i, ByVal s, 2
   'ElseIf Len(s) = 1 Then
   '   CopyMemory i, ByVal s, 1
   End If
   
   CVI = i
End Function
Public Function CVL(s As String) As Long
   Dim i As Long
   
   If Len(s) <> 4 Then
      Err.Raise 1000, "CVL", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 4
   End If
   
   CVL = i
End Function
Public Function CVD(s As String) As Double
   Dim i As Double
   
   If Len(s) <> 8 Then
      Err.Raise 1000, "CVD", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 8
   End If
   
   CVD = i
End Function
Public Function CVS(s As String) As Single
   Dim i As Single
   
   If Len(s) <> 4 Then
      Err.Raise 1000, "CVS", "Invalid string argument"
   Else
      CopyMemory i, ByVal s, 4
   End If
   
   CVS = i
End Function
Public Function MKI(ByVal i As Integer) As String
    Dim s As String
    
    s = String(2, 0)
    CopyMemory ByVal s, i, 2
    
    MKI = s
End Function

Public Function MKL(ByVal i As Long) As String
    Dim s As String
    
    s = String(4, 0)
    CopyMemory ByVal s, i, 4
    
    MKL = s
End Function
Public Function MKS(ByVal i As Double) As String
    Dim s As String
    
    s = String(4, 0)
    CopyMemory ByVal s, i, 4
    
    MKS = s
End Function

Public Function MKD(ByVal i As Double) As String
    Dim s As String
    
    s = String(8, 0)
    CopyMemory ByVal s, i, 8
    
    MKD = s
End Function

