VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   Caption         =   "Form3"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form3"
   ScaleHeight     =   242
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Disable"
      Height          =   255
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Disable"
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Disable"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Testing"
      Height          =   855
      Left            =   2400
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   855
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Testing"
      Height          =   855
      Left            =   1200
      Picture         =   "Form3.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Testing"
      Height          =   855
      Left            =   0
      Picture         =   "Form3.frx":2AAC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   840
      Picture         =   "Form3.frx":524E
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   9600
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents FBsp2 As CustomButtons
Attribute FBsp2.VB_VarHelpID = -1

Private Sub Command1_Click()
Debug.Print Command1.Width
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If Command1.Enabled Then
Command1.Enabled = False
Command5.Caption = "Enable"
Else
Command1.Enabled = True
Command5.Caption = "Disable"
End If
End Sub

Private Sub Command6_Click()
If Command2.Enabled Then
Command2.Enabled = False
Command6.Caption = "Enable"
Else
Command2.Enabled = True
Command6.Caption = "Disable"
End If
End Sub

Private Sub Command7_Click()
If Command4.Enabled Then
Command4.Enabled = False
Command7.Caption = "Enable"
Else
Command4.Enabled = True
Command7.Caption = "Disable"
End If
End Sub

Private Sub FBsp2_Resizing(hwnd As Long, focus As Boolean, resizetype As Long, syscolchange As Boolean)
BitBltt Me.hdc, 0, 0, ScaleWidth, Command1.Height + 3, Picture1.hdc, 0, 0, vbSrcCopy
If ScaleWidth > Picture1.Width Then
BitBltt Me.hdc, Picture1.Width, 0, ScaleWidth - Picture1.Width, Command1.Height + 3, Picture1.hdc, 0, 0, vbSrcCopy
End If

Line (0, Command1.Height + 3)-(ScaleWidth, Command1.Height + 3), QBColor(8)
Line (0, Command1.Height + 4)-(ScaleWidth, Command1.Height + 4), QBColor(15)
Command3.Left = ScaleWidth - Command3.Width
FBsp2.RefreshButtons Command3
End Sub

Private Sub Form_Load()
'Me.Picture = Picture1
BitBltt Me.hdc, 0, 0, ScaleWidth, Command1.Height + 3, Picture1.hdc, 0, 0, vbSrcCopy
Line (0, Command1.Height + 3)-(ScaleWidth, Command1.Height + 3), QBColor(8)
Line (0, Command1.Height + 4)-(ScaleWidth, Command1.Height + 4), QBColor(15)
Set FBsp2 = New CustomButtons
FBsp2.MakeCustomButtons Me
FBsp2.TransparentButton = True
FBsp2.HoverStyle = True
FBsp2.RefreshButtons
End Sub

Private Sub Form_Unload(Cancel As Integer)
FBsp2.HoverStyle = False
End Sub
