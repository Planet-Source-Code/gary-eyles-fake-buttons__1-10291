VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Buttons.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
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
      Height          =   210
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "SYSbutton"
      Top             =   0
      Width           =   240
   End
   Begin VB.CommandButton Command8 
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
      Height          =   210
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "SYSbutton"
      Top             =   0
      Width           =   240
   End
   Begin VB.CommandButton Command7 
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
      Height          =   210
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "SYSbutton"
      Top             =   0
      Width           =   240
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command1"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Testing"
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4680
      TabIndex        =   7
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
'Command1.Caption = "Testing Testing Testing"
End Sub

Private Sub Command2_Click()
'Dim tmp As String
'tmp = FindControl(Command9.hwnd)
'Command2.Caption = tmp

Dim tmppict As StdPicture
Set tmppict = Me.Icon
'Me.Picture = tmppict
'Me.PaintPicture tmppict, 100, 0, 20, 20, 0, 0, 10, 10, vbSrcCopy
'tmppict.Render Me.hdc, 0, 0, 20, 20, 0, 0, 30, 30, 1
'Me.Picture = Me.Icon
Me.Refresh
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click()
If Me.WindowState = 0 Then
Me.WindowState = 2
Command7.Caption = "2"
Else
Me.WindowState = 0
Command7.Caption = "1"
End If
End Sub

Private Sub Command8_Click()
Form1.WindowState = 1
End Sub

Private Sub Command9_Click()
End
End Sub

Private Sub Form_Load()
AttachMessage Me, Me.hwnd, WM_DRAWITEM
ShowTitleBar False

Command7.Top = 2
Command8.Top = 2
Command9.Top = 2
Command9.Left = Command9.Left - 2
Picture1.Height = Command7.Height + 4
'Picture1.Refresh
End Sub

Private Sub Form_Resize()
'Line (0, 0)-(ScaleWidth - 200, Command9.Top + Command9.Height + 1), QBColor(1), BF
Command9.Left = ScaleWidth - Command9.Width - 2
Command7.Left = Command9.Left - Command7.Width - 2
Command8.Left = Command7.Left - Command8.Width

Picture1.FontBold = True
Picture1.CurrentX = 30
Picture1.CurrentY = 30
Picture1.ForeColor = QBColor(15)
Picture1.Print "Custom Buttons"
'Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
DetachMessage Me, Me.hwnd, WM_DRAWITEM
End Sub



Private Sub Picture1_DblClick()
If Me.WindowState = 0 Then
Me.WindowState = 2
Command7.Caption = "2"
Else
Me.WindowState = 0
Command7.Caption = "1"
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Form1.hwnd, WM_NCLBUTTONDOWN, _
HTCAPTION, 0&)
End If
End Sub
