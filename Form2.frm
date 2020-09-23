VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form2"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1815
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dithered Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Shadow Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Hilight Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Focus Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000C0&
      Caption         =   "Disabled Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton CloseButton 
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "SYSbutton"
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000080&
      Caption         =   "Text Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000040&
      Caption         =   "Back Colour"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FBs As CustomButtons
Attribute FBs.VB_VarHelpID = -1
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWW_HPARENT = (-8)

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Form1.ChangeColour 0
End Sub

Private Sub Command2_Click()
Form1.ChangeColour ForeColour
End Sub

Private Sub Command3_Click()
Form1.ChangeColour DisabledTextColour
End Sub

Private Sub Command4_Click()
Form1.ChangeColour FocusForeColour
End Sub

Private Sub Command5_Click()
Form1.ChangeColour BorderHilightColour
End Sub

Private Sub Command6_Click()
Form1.ChangeColour BorderShadowColour
End Sub

Private Sub Command7_Click()
Form1.ChangeColour DitheredColour
End Sub

Private Sub Form_Load()
Set FBs = New CustomButtons

'FBs.HoverStyle = True
FBs.MakeCustomButtons Form2, False
FBs.MakeTitleBar Picture1
FBs.TextColour = QBColor(13)
FBs.FocusTextColour = QBColor(10)

Picture1.Height = 15
CloseButton.Top = 2
CloseButton.Width = Picture1.Height - 4
CloseButton.Height = Picture1.Height - 4
CloseButton.Left = Width / 15 - Picture1.Height - 4
Picture1.FontBold = True

Form2.Height = Command7.Top * 15 + Command7.Height * 15 + Picture1.Height * 15
'FBs.SetMinMaxInfo 100, Picture1.Height * 5, 400, 300
FBs.SetMinMaxInfo Width / 15, Height / 15, Width / 15, Height / 15

Dim iRval As Integer
'    If (bState) Then
        SetWindowLong Me.hwnd, SWW_HPARENT, Form1.hwnd
'    Else
'        SetWindowLong frmTopMost.hwnd, SWW_HPARENT, 0
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
FBs.UnMakeCustomButtons Form2
SetWindowLong Me.hwnd, SWW_HPARENT, 0
End Sub

Public Sub FBs_Resizing(hwnd As Long, focus As Boolean, resizetype As Long, syscolchange As Boolean)
Picture1.Width = Width / 15
CloseButton.Left = Width / 15 - Picture1.Height - 4

If hwnd = -1 And focus Then
    Picture1.BackColor = QBColor(12)
    Picture1.ForeColor = QBColor(15)
ElseIf hwnd = -1 And Not focus Then
    Picture1.BackColor = QBColor(8)
    Picture1.ForeColor = QBColor(7)
End If

If hwnd = -1 Then
    FBs.cDrawText Picture1, "Toolbar", 2, CloseButton.Left - 4, 2, Picture1.Height
    Picture1.Refresh
    Exit Sub
End If

If focus Or hwnd = Form1.hwnd Then
    Picture1.BackColor = QBColor(12)
    Picture1.ForeColor = QBColor(15)
    If hwnd <> Form1.hwnd Then
        Form1.FakeButtons_Resizing Form1.hwnd, True, Form1.WindowState, False
    End If
Else
    Picture1.BackColor = QBColor(8)
    Picture1.ForeColor = QBColor(7)
    Form1.FakeButtons_Resizing Form1.hwnd, False, Form1.WindowState, False
End If

FBs.cDrawText Picture1, "Toolbar", 2, CloseButton.Left - 4, 2, Picture1.Height
Picture1.Refresh
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    FBs.ClickTitleBar
End If
End Sub
