VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Fake Buttons"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "Test2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Test2.frx":0ECA
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   Begin VB.CommandButton Command31 
      Caption         =   "Show Form3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1200
      Width           =   1215
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   480
      Picture         =   "Test2.frx":2670C
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   36
      Top             =   5640
      Width           =   3975
      Begin VB.CommandButton Command30 
         Height          =   615
         Index           =   2
         Left            =   1800
         Picture         =   "Test2.frx":610CE
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command30 
         Height          =   615
         Index           =   1
         Left            =   960
         Picture         =   "Test2.frx":613D8
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton Command30 
         Height          =   615
         Index           =   0
         Left            =   120
         Picture         =   "Test2.frx":616E2
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Testing"
      Height          =   1335
      Left            =   4920
      Picture         =   "Test2.frx":619EC
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Own Backcolor (TRUE)"
      Height          =   855
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H0080C0FF&
      Caption         =   "Old Graphical style"
      Height          =   855
      Left            =   1680
      Picture         =   "Test2.frx":628B6
      Style           =   1  'Graphical
      TabIndex        =   29
      Tag             =   "normal"
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "SYSbutton"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "SYSbutton"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   6.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "SYSbutton"
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   0
      Width           =   5400
   End
   Begin VB.CommandButton Command11 
      Caption         =   "&Transparent"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      Picture         =   "Test2.frx":63860
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      Picture         =   "Test2.frx":64AB2
      ScaleHeight     =   255
      ScaleWidth      =   1350
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000FF00&
      Caption         =   "Indexed"
      Height          =   375
      Index           =   2
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Test4"
      Height          =   735
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Show Form2"
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      Caption         =   "&Hover Style"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Resize Buttons"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Different Colours"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Disabled"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Original Colours"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Old button"
      Height          =   615
      Left            =   3360
      TabIndex        =   19
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Draw &Focus Rect (OFF)"
      Height          =   735
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command21 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command22 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   800
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Test3"
      Height          =   735
      Left            =   2040
      Picture         =   "Test2.frx":65D04
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Test2"
      Height          =   735
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Test Test Test Test"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0000C000&
      Caption         =   "Indexed"
      Height          =   375
      Index           =   1
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00008000&
      Caption         =   "Indexed"
      Height          =   375
      Index           =   0
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Max/Restore"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Minimize"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command25 
      Caption         =   "&Proper size"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Old Style (FALSE)"
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Having transparent buttons requires the parent of the button to have its autoredraw set to TRUE."
      Height          =   735
      Left            =   1560
      TabIndex        =   12
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "18"
      Height          =   195
      Left            =   5280
      TabIndex        =   28
      Top             =   1080
      Width           =   180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long

Dim Resized As Boolean
Private WithEvents FakeButtons As CustomButtons
Attribute FakeButtons.VB_VarHelpID = -1
Private WithEvents FBsp As CustomButtons
Attribute FBsp.VB_VarHelpID = -1

Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
Private Type ChooseColor
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long
'Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long
'Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Dim OFName As OPENFILENAME
Dim CustomColors() As Byte

Private Function ShowColour(Optional col As Long = 0) As Long
    Dim cc As ChooseColor
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hWndOwner = Me.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = cdlCCRGBInit Or cdlCCFullOpen

    cc.rgbResult = col
    'Show the 'Select Color'-dialog
    If ChooseColor(cc) <> 0 Then
        ShowColour = cc.rgbResult
        'CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColour = -1
    End If
End Function

Sub DD(hwnd As Long, hdc As Long, pict As Long, xtmp, ytmp)
Module2.DrawDithered hwnd, hdc, pict, xtmp, ytmp
End Sub

Sub ChangeColour(num As ChangeCol)
Dim CmDlg As New cCommonDialog
'CmDlg.HookDialog = False
CmDlg.hwnd = Me.hwnd

If num = BackColour Then
    CmDlg.Color = FakeButtons.BackColour
ElseIf num = FocusForeColour Then
    CmDlg.Color = FakeButtons.FocusTextColour
ElseIf num = ForeColour Then
    CmDlg.Color = FakeButtons.TextColour
ElseIf num = DisabledTextColour Then
    CmDlg.Color = FakeButtons.DisabledTextColour
ElseIf num = BorderHilightColour Then
    CmDlg.Color = FakeButtons.BorderHilightColour
ElseIf num = BorderShadowColour Then
    CmDlg.Color = FakeButtons.BorderShadowColour
ElseIf num = DitheredColour Then
    CmDlg.Color = FakeButtons.ShadeColor
End If

CmDlg.flags = cdlCCRGBInit Or cdlCCFullOpen
CmDlg.ShowColor

If num = BackColour Then
    FakeButtons.BackColour = CmDlg.Color
ElseIf num = FocusForeColour Then
    FakeButtons.FocusTextColour = CmDlg.Color
ElseIf num = ForeColour Then
    FakeButtons.TextColour = CmDlg.Color
ElseIf num = DisabledTextColour Then
    FakeButtons.DisabledTextColour = CmDlg.Color
ElseIf num = BorderHilightColour Then
    FakeButtons.BorderHilightColour = CmDlg.Color
ElseIf num = BorderShadowColour Then
    FakeButtons.BorderShadowColour = CmDlg.Color
ElseIf num = DitheredColour Then
    FakeButtons.ShadeColor = CmDlg.Color
End If

'Unload CmDlg
Set CmDlg = Nothing
FakeButtons.RefreshButtons
End Sub


Private Sub DrawTitleBar()
Picture1.Height = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 1
Label2.Caption = Picture1.Height

Command3.Top = 2
Command5.Top = 2
Command6.Top = 2
Command3.Height = Picture1.Height - 4
Command5.Height = Picture1.Height - 4
Command6.Height = Picture1.Height - 4
Command3.Width = Command3.Height + 2
Command5.Width = Command5.Height + 2
Command6.Width = Command6.Height + 2

'Command3.Visible = False
'Command5.Visible = False
'Command6.Visible = False

'Command3.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 12
'Command5.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 11
'Command6.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 12

'Newer
'Command3.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 13
'Command5.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 13
'Command6.FontSize = FakeButtons.GetSysMetrics(SM_CYCAPTION) - 13
End Sub

Private Sub Command1_Click()
FakeButtons.BackColour = RGB(100, 200, 100)
FakeButtons.TextColour = RGB(0, 0, 0)
FakeButtons.BorderHilightColour = QBColor(14)
FakeButtons.BorderShadowColour = QBColor(4)
FakeButtons.FocusTextColour = RGB(0, 0, 255)
FakeButtons.DisabledTextColour = QBColor(5)
'FakeButtons.RepaintWindow
FakeButtons.RefreshButtons
End Sub

Private Sub Command10_Click()
Exit Sub

Dim Gh As Long
Gh = CreateSolidBrush(QBColor(2))

DrawState Form1.hdc, 0, 0, Form1.Command27.Picture, 0, _
    Form1.Command1.Left + Form1.Command1.Width, Form1.Picture1.Height, _
    Form1.Command27.Picture.Width, Form1.Command27.Picture.Height, _
    DST_BITMAP 'Or DSS_MONO Or DSS_UNION
DrawState Form1.hdc, Gh, 0, Form1.Command27.Picture, 0, _
    Form1.Command1.Left + Form1.Command1.Width, Form1.Picture1.Height, _
    Form1.Command27.Picture.Width, Form1.Command27.Picture.Height, _
    DST_BITMAP Or DSS_MONO Or DSS_UNION

DeleteObject Gh
'Public Const DST_COMPLEX = &H0&
'Public Const DST_TEXT = &H1&
'Public Const DST_PREFIXTEXT = &H2&
'Public Const DST_ICON = &H3&
'Public Const DST_BITMAP = &H4&

'Public Const DSS_NORMAL = &H0&
'Public Const DSS_UNION = &H10& ' Dither
'Public Const DSS_DISABLED = &H20&
'Public Const DSS_MONO = &H80& ' Draw in colour of brush specified in hBrush
'Public Const DSS_RIGHT = &H8000&
Me.Refresh
End Sub

Private Sub Command11_Click()
If FakeButtons.TransparentButton Then
    FakeButtons.TransparentButton = False
Else
    FakeButtons.TransparentButton = True
End If
'FakeButtons.RepaintWindow
FakeButtons.RefreshButtons
End Sub

Private Sub Command13_Click()
If FakeButtons.DrawFocusRct Then
FakeButtons.DrawFocusRct = False
Command13.Caption = "Draw Focus Rect (OFF)"
Else
FakeButtons.DrawFocusRct = True
Command13.Caption = "Draw Focus Rect (ON)"
End If
End Sub

Private Sub Command15_Click()
On Error Resume Next
Dim TmpButton As Object

'Reduce the width of each button.
'This is becuase of the snap to grid,
'and reducing the width of the buttons
'by one looks better to me.
If Resized = False Then
    For Each TmpButton In Form1
    If TmpButton.Tag <> "SYSbutton" And _
            TypeOf TmpButton Is CommandButton Then
        TmpButton.Width = TmpButton.Width - 1
        TmpButton.Height = TmpButton.Height - 1
    End If
    Next TmpButton
    Resized = True
Else
    For Each TmpButton In Form1
    If TmpButton.Tag <> "SYSbutton" And _
            TypeOf TmpButton Is CommandButton Then
        TmpButton.Width = TmpButton.Width + 1
        TmpButton.Height = TmpButton.Height + 1
    End If
    Next TmpButton
    Resized = False
End If
End Sub

Private Sub Command16_Click()
If FakeButtons.HoverStyle = False Then
    FakeButtons.HoverStyle = True
'    FBsp.HoverStyle = True
Else
    FakeButtons.HoverStyle = False
'    FBsp.HoverStyle = False
End If
FakeButtons.RefreshButtons
End Sub

Private Sub Command17_Click()
Form2.Show
End Sub

Private Sub Command18_Click()
If Command3.Enabled Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If
End Sub

Private Sub Command19_Click()
If Command5.Enabled Then
    Command5.Enabled = False
Else
    Command5.Enabled = True
End If
End Sub

Private Sub Command20_Click()
If Command6.Enabled Then
    Command6.Enabled = False
Else
    Command6.Enabled = True
End If
End Sub

Private Sub Command21_Click()
'FakeButtons_Resizing(focus As Boolean, resizetype As Long)
Picture1.Height = Picture1.Height + 1
Command3.Height = Picture1.Height - 4
Command5.Height = Picture1.Height - 4
Command6.Height = Picture1.Height - 4
Command3.Width = Command3.Height + 2
Command5.Width = Command5.Height + 2
Command6.Width = Command6.Height + 2
Label2.Caption = Picture1.Height
FakeButtons_Resizing Me.hwnd, True, Me.WindowState, False
End Sub

Private Sub Command22_Click()
'FakeButtons_Resizing(focus As Boolean, resizetype As Long)
If Picture1.Height = 15 Then Exit Sub
Picture1.Height = Picture1.Height - 1
Command3.Height = Picture1.Height - 4
Command5.Height = Picture1.Height - 4
Command6.Height = Picture1.Height - 4
Command3.Width = Command3.Height + 2
Command5.Width = Command5.Height + 2
Command6.Width = Command6.Height + 2
Label2.Caption = Picture1.Height
FakeButtons_Resizing Me.hwnd, True, Me.WindowState, False
End Sub

Private Sub Command24_Click()
If FakeButtons.OldStyleSystemButtons Then
    FakeButtons.OldStyleSystemButtons = False
    Command24.Caption = "Old Style (FALSE)"
Else
    FakeButtons.OldStyleSystemButtons = True
    Command24.Caption = "Old Style (TRUE)"
End If
FakeButtons.RefreshButtons
End Sub

Private Sub Command25_Click()
Picture1.Height = FakeButtons.GetSysMetrics(SM_CYCAPTION)
Picture1.Height = Picture1.Height - 1
Command3.Height = Picture1.Height - 4
Command5.Height = Picture1.Height - 4
Command6.Height = Picture1.Height - 4
Command3.Width = Command3.Height + 2
Command5.Width = Command5.Height + 2
Command6.Width = Command6.Height + 2
Label2.Caption = Picture1.Height
FakeButtons_Resizing Me.hwnd, True, Me.WindowState, False
End Sub

Private Sub Command26_Click()
If FakeButtons.OwnBackColour Then
    FakeButtons.OwnBackColour = False
    Command26.Caption = "Own Backcolor (FALSE)"
Else
    FakeButtons.OwnBackColour = True
    Command26.Caption = "Own Backcolor (TRUE)"
End If
FakeButtons.RefreshButtons
End Sub

Private Sub Command27_Click()
'Debug.Print Int(Command27.Picture.Width / 26)
'Debug.Print Command27.Picture.Height
'Dim gaztmp As BITMAP
'gaztmp = Space(20)
'CopyMemory gaztmp, Command27.Picture.handle, Len(gaztmp)
'Debug.Print gaztmp.bmWidthBytes
'Command27.Caption = gaztmp.bmBits
End Sub

Private Sub Command28_Click()
If Command27.Enabled Then
    Command27.Enabled = False
    Command28.Caption = "Disabled"
Else
    Command27.Enabled = True
    Command28.Caption = "Enabled"
End If
'FakeButtons.RefreshButtons
End Sub

Private Sub Command29_Click()
If Command9.Enabled Then
Command9.Enabled = False
Command29.Caption = "Disabled"
Else
Command9.Enabled = True
Command29.Caption = "Enabled"
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command31_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
FakeButtons.BackColour = FakeButtons.SysColour(COLOR_BTNFACE)
FakeButtons.TextColour = FakeButtons.SysColour(COLOR_BTNTEXT)
FakeButtons.BorderHilightColour = FakeButtons.SysColour(COLOR_BTNHIGHLIGHT)
FakeButtons.BorderShadowColour = FakeButtons.SysColour(COLOR_BTNSHADOW)
FakeButtons.FocusTextColour = QBColor(4)
FakeButtons.DisabledTextColour = QBColor(8)
'FakeButtons.RepaintWindow
FakeButtons.RefreshButtons
End Sub

Private Sub Command5_Click()
If Me.WindowState = 0 Then
    FakeButtons.WindowsCommands Me, SW_MAXIMIZE
ElseIf Me.WindowState = 2 Then
    FakeButtons.WindowsCommands Me, SW_RESTORE
End If
End Sub

Private Sub Command6_Click()
FakeButtons.WindowsCommands Me, SW_MINIMIZE
End Sub

Private Sub Command7_Click()
'Dim r
'r = ShowColour
End Sub

Private Sub Command9_Click()
'Module2.DrawDithered Me.hwnd, Me.hdc, Command27.Picture, Command1.Left + Command1.Width, Picture1.Height
'Me.Refresh

'Debug.Print "1) " & Command8.Picture
'Debug.Print "2) " & Command9.Picture
'Debug.Print "3) " & Command9.Picture.Handle
'Debug.Print "Icon " & Me.Icon

'Set m_pic = Command27.Picture
'Module2.pbGetBitmapIntoDC
'BitBlt Me.hdc, 0, Picture1.Height, m_lBitmapW, m_lBitmapH, m_lHdc, 0, 0, vbSrcCopy
'Dim trTmp As RECT
'trTmp.Top = 0
'trTmp.Left = 0
'trTmp.Right = m_lBitmapW - 1
'trTmp.Bottom = m_lBitmapH - 1
'TransBlt Me.hdc, Me.hdc, m_lHdc, trTmp, 0, Picture1.Height, 0
'Me.Refresh
'Module2.pClearUp
End Sub

Private Sub Form_Load()
On Error Resume Next
'Center the form, I know there is an option
'for this on VB6 but it centers the form
'every time I refresh the buttons.
'This bit of code doesn't.
Me.Move Screen.Width / 2 - (Width / 2), _
        Screen.Height / 2 - (Height / 2), _
        Width, Height - 36 * 15

Set FakeButtons = New CustomButtons
Set FBsp = New CustomButtons
FBsp.BackColour = QBColor(7)
'FBsp.OwnBackColour = False
FBsp.TransparentButton = True
FBsp.MakeCustomButtons Picture4

Call DrawTitleBar

'This makes the new buttons, don't forget
'the style needs to be set to graphical
'for any effect.
FakeButtons.MakeCustomButtons Form1, False
'The picturebox can be dragged like the
'titlebar.
FakeButtons.MakeTitleBar Picture1
Command15_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Unload the custom buttons other wise
'your likely to crash.
FakeButtons.UnMakeTitleBar Picture1
FakeButtons.UnMakeCustomButtons Form1
Unload Form2
End Sub

Public Sub FakeButtons_Resizing(hwnd As Long, focus As Boolean, resizetype As Long, syscolchange As Boolean)
'focus returns whether the window is
'in focus or not.

'resizetype returns whether the window is:
'0=Normal
'1=Minimised
'2=Maximised

'sysColChange tells you whether there is a
'system wide colour change.

'Change the button to display either
'restore or maximise.
If resizetype = 0 Then
    Command5.Caption = "1"
ElseIf resizetype = 2 Then
    Command5.Caption = "2"
End If

If syscolchange Then
    Call DrawTitleBar
End If

If focus And hwnd <> Form2.hwnd Then
    Form2.FBs_Resizing -1, True, Form2.hwnd, False
ElseIf Not focus And hwnd <> Form2.hwnd Then
    Form2.FBs_Resizing -1, False, Form2.hwnd, False
End If

'Draw and position the title bar. COOL!
With Picture1
.Width = ScaleWidth
.AutoRedraw = True

Dim xx As Long
Dim col As Long
For xx = 0 To Picture1.Width Step 5

If focus Or hwnd = Form2.hwnd Then
    col = (255 / Picture1.Width) * xx
    col = RGB(col, 0, 0)
    'Form2.FBs_Resizing -1, True, Form2.WindowState, False
Else
    col = (155 / Picture1.Width) * xx
    col = RGB(col, col, col)
    'Form2.FBs_Resizing Form2.hwnd, False, Form2.WindowState, False
End If

Picture1.Line (xx, 0)-(xx + 5, Picture1.Height), col, BF
Next

.ForeColor = QBColor(15)
.FontBold = True
FakeButtons.cDrawText Picture1, Me.Caption, _
        Picture1.Height + 2, _
        Command6.Left - 3, _
        0, Picture1.Height

If ScaleWidth - 80 > Picture1.TextWidth(Me.Caption) + Picture2.Width + 4 Then
Picture1.PaintPicture Picture3, _
    Picture1.TextWidth(Me.Caption) + Picture1.Height + 4, 2, _
    Picture2.Width, Picture1.Height - 4, 0, 0, _
    Picture2.Width, Picture2.Height, vbSrcAnd
Picture1.PaintPicture Picture2, _
    Picture1.TextWidth(Me.Caption) + Picture1.Height + 4, 2, _
    Picture2.Width, Picture1.Height - 4, 0, 0, _
    Picture2.Width, Picture2.Height, vbSrcPaint
End If

FakeButtons.cDrawIcon Picture1, Me, 2, 2, _
    Picture1.Height - 2, Picture1.Height - 4, di_normal

'FakeButtons.DrawSystemButtons Me.hwnd
.Refresh
.AutoRedraw = False
End With

'6,5,3
Command3.Left = ScaleWidth - Command3.Width - 2
Command5.Left = Command3.Left - Command5.Width - 2
Command6.Left = Command5.Left - Command6.Width

FakeButtons.SetMinMaxInfo _
 (Command3.Width * 4) + 50, _
 Picture1.Height + 6, Screen.Width / 15 + 20, _
 Screen.Height / 15
 
'Command11.Top = ScaleHeight - Command11.Height - 10
FakeButtons.RefreshButtons Command11
End Sub

Private Sub Picture1_DblClick()
Dim mc As POINTAPI
GetCursorPos mc
mc.x = mc.x - (Me.Left / 15) - FakeButtons.GetSysMetrics(SM_CXFRAME) + 1
'Debug.Print mc.X
If mc.x > Picture1.Height Then
    If mc.x < Command6.Left - 5 Then Command5_Click
Else
    Unload Me
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And x < Command6.Left - 2 And Me.WindowState = 0 Then
    FakeButtons.ClickTitleBar
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Xt As Long
Dim Yt As Long
Dim rtmp As RECT
GetWindowRect Me.hwnd, rtmp
Xt = x + rtmp.Left
Yt = y + rtmp.Top

If Button = 2 And x < Picture1.Height Then
    'SendKeys "% "
    GetWindowRect Picture1.hwnd, rtmp
    Xt = rtmp.Left
    Yt = rtmp.Bottom
    FakeButtons.SystemMenu Me, Xt, Yt
ElseIf Button = 2 And x > Picture1.Height And x < Command6.Left - 5 Then
    GetWindowRect Me.hwnd, rtmp
    Xt = x + rtmp.Left
    Yt = y + rtmp.Top
    FakeButtons.SystemMenu Me, Xt + 3, Yt + 3
End If
End Sub

