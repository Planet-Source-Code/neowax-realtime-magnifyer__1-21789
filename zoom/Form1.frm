VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Value           =   1  'Aktiviert
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Text            =   "100"
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Text            =   "100"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3495
      LargeChange     =   10
      Left            =   120
      Max             =   1000
      Min             =   1
      MousePointer    =   7  'Größenänderung N S
      TabIndex        =   2
      Top             =   960
      Value           =   100
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   960
      Max             =   1000
      Min             =   2
      MousePointer    =   9  'Größenänderung W O
      TabIndex        =   1
      Top             =   120
      Value           =   100
      Width           =   5535
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6000
      Top             =   3960
   End
   Begin VB.PictureBox PictureBox1 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   600
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   389
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' *************************************************
' realtime magnifyer
' the code was written by neowax
' you may use it for your own application
' i wrote this, cause there were no really good magnifyer
' the advantage: it used only 2-6% of the system-resources
' while moving and zooming
' for further information feel free to
' contact me under neowax@uni.de
' please have patience while receiving answer
' *************************************************

Dim dhwnd As Long, dhdc As Long
Dim x As Integer, y As Integer
Dim w As Integer, h As Integer
Dim sw As Integer, sh As Integer
Dim zoom As Integer
Dim mouse As PointAPI

Private Sub Command1_Click()
Text1.Text = "100"
Text2.Text = "100"
VScroll1.Value = 100
HScroll1.Value = 100
End Sub

Private Sub Form_Load()
dhwnd = GetDesktopWindow ' get desktop window
dhdc = GetDC(dhwnd)      ' get display device
End Sub

Private Sub Form_Resize()
PictureBox1.Width = Form1.ScaleWidth - PictureBox1.Left
PictureBox1.Height = Form1.ScaleHeight - PictureBox1.Top
HScroll1.Width = Form1.ScaleWidth - HScroll1.Left
VScroll1.Height = Form1.ScaleHeight - VScroll1.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ReleaseDC(dhwnd, dhdc) ' it's important you free tha dc, cause windows may crash
End Sub

Private Sub HScroll1_Scroll()
Text1.Text = HScroll1.Value
If Check1.Value = 1 Then Text2.Text = HScroll1.Value: VScroll1.Value = HScroll1.Value
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_Change()
If Check1.Value = 1 Then Text2.Text = Text1.Text
End Sub

Private Sub Text2_Change()
If Check1.Value = 1 Then Text1.Text = Text2.Text
End Sub


Private Sub Timer1_Timer()
GetCursorPos mouse                                  ' capture mouse-position
Me.Caption = "X: " & mouse.x & ", Y: " & mouse.y    ' write position 2 window-title
w = PictureBox1.ScaleWidth                          ' destination width
h = PictureBox1.ScaleHeight                         ' destination height
sw = w * (1 / (IIf(Text1.Text = "" Or Text1.Text < 2, 2, Text1.Text) / 100)) ' source width
sh = h * (1 / (IIf(Text2.Text = "", 1, Text2.Text) / 100))                   ' source height
x = mouse.x - sw \ 2                                ' x source position (center to destination)
y = mouse.y - sh \ 2                                ' y source position (center to destination)
PictureBox1.Cls                                     ' clean picturebox
StretchBlt PictureBox1.hDC, 0, 0, w, h, dhdc, x, y, sw, sh, SRCCOPY  ' copy desktop (source) and strech to picturebox (destination)
End Sub

Private Sub VScroll1_Scroll()
Text2.Text = VScroll1.Value
If Check1.Value = 1 Then Text1.Text = VScroll1.Value: HScroll1.Value = HScroll1.Value
End Sub
