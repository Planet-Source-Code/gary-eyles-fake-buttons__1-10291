VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFormBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass
Private m_emr As EMsgResponse

Private m_hWnd As Long
Private m_oBackColor As OLE_COLOR
Private m_cT As cTile

Private Const GW_CHILD = 5

Private Const WM_PAINT = &HF
Private Const WM_SIZE = &H5
Private Const WM_ERASEBKGND = &H14

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Gets/sets the back colour to draw in a form when no bitmap has been set to tile into the backdrop."
    BackColor = m_oBackColor
End Property
Public Property Let BackColor(ByVal oColor As OLE_COLOR)
    m_oBackColor = oColor
End Property

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function
Private Sub PaintBackground(ByVal hParent As Long, ByVal oColor As OLE_COLOR, ByVal nPenStyle As Long)

Dim tPS As PAINTSTRUCT  ' PaintStruct used in painting
Dim hNewBrush As Long   ' handle to new brush
Dim nRectangle As RECT  ' rectangle structure
Dim lColor As Long      ' GDI version of OLE_COLOR to draw in

    
    ' Don't draw when we are iconised:
    If Not (IsIconic(m_hWnd)) Then

        ' We do this after any other messages have done their stuff
        m_emr = emrPostProcess
                
        ' Begin painting:
        BeginPaint m_hWnd, tPS
        
        'Retrieve the rectangle size that describes the client area of the 'window
        GetClientRect m_hWnd, nRectangle
        
        ' Are we using a bitmap?
        If (m_cT.Filename = "") Then
            ' If not, just colour in with the selected backcolour:
            lColor = TranslateColor(m_oBackColor)
            'Create a solid color brush; return handle
            hNewBrush = CreateSolidBrush(lColor)
            ' Fill the client area
            FillRect tPS.hdc, nRectangle, hNewBrush
            'Delete brush
            DeleteObject hNewBrush
        Else
            ' Tile the bitmap into the background:
            m_cT.TileArea tPS.hdc, nRectangle.Left, nRectangle.TOp, nRectangle.Right - nRectangle.Left, nRectangle.Bottom - nRectangle.TOp
        End If
            
        ' End Painting:
        EndPaint m_hWnd, tPS
    End If
    
End Sub
Property Get Tile() As cTile
Attribute Tile.VB_Description = "Returns a reference to the bitmap tiling object used to draw the tiled bitmap into the object."
    ' Provide access to the MDI back tiling object:
    Set Tile = m_cT
End Property
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the bitmap background.  Call this if you change the background whilst the form is visible."
Dim tR As RECT
Dim tP As POINTAPI
    If (m_hWnd <> 0) Then
        GetWindowRect m_hWnd, tR
        tP.x = tR.Left
        tP.y = tR.TOp
        ScreenToClient m_hWnd, tP
        tR.Left = tP.x
        tR.TOp = tP.y
        tP.x = tR.Right
        tP.y = tR.Bottom
        ScreenToClient m_hWnd, tP
        tR.Right = tP.x
        tR.Bottom = tP.y
        InvalidateRect m_hWnd, tR, 1
    End If
End Sub

Private Sub Class_Initialize()
    Set m_cT = New cTile
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (iMsg = WM_PAINT) Then
        PaintBackground m_hWnd, m_oBackColor, PS_SOLID
    End If
End Function
Public Sub Init(ByRef frmThis As Object)
Attribute Init.VB_Description = "Sets the object that the will be tiled with a bitmap and commences drawing the background."
    
    ' Ensure we don't have an existing link:
    Destroy
    
    ' Store the window used to draw the background of the MDI form:
    If TypeOf frmThis Is MDIForm Then
        m_hWnd = GetWindow(frmThis.hwnd, GW_CHILD)
    Else
        m_hWnd = frmThis.hwnd
    End If
    ' Store the current back colour:
    m_oBackColor = frmThis.BackColor
    
    ' Start subclassing
    AttachMessage Me, m_hWnd, WM_PAINT
    
End Sub
Public Sub Destroy()
Attribute Destroy.VB_Description = "Clears up any resources associated with applying a bitmap to a form.  Call this before your application terminates."
    
    ' If we have a hWnd, then remove the subclassed messages:
    If (m_hWnd <> 0) Then
        DetachMessage Me, m_hWnd, WM_PAINT
        m_hWnd = 0
    End If
    
End Sub
Private Sub Class_Terminate()
    ' Clear up:
    Set m_cT = Nothing
    Destroy
End Sub




