VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000016&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Internet Explorer"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2190
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   146
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTray 
      BackColor       =   &H80000016&
      Caption         =   "Start in system tray"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4335
      TabIndex        =   16
      Top             =   975
      Width           =   1935
   End
   Begin VB.VScrollBar scrSecs 
      Height          =   300
      Left            =   1800
      Max             =   1
      Min             =   60
      TabIndex        =   6
      Top             =   1740
      Value           =   1
      Width           =   180
   End
   Begin VB.TextBox txtSecs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1770
      Width           =   375
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1200
      Top             =   -7500
   End
   Begin VB.CheckBox chkKeepAwake 
      BackColor       =   &H80000016&
      Caption         =   "Enable Keep Alive"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Value           =   2  'Grayed
      Width           =   1935
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   17
      ToolTipText     =   " Minimize to tray "
      Top             =   2160
      Width           =   165
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "v2.0 made by Frisco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4320
      TabIndex        =   15
      Top             =   600
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "View v1.0 source code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   180
      Left            =   4320
      MouseIcon       =   "frmMain.frx":1BB2
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   300
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Found on Planet Source Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "2.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   150
      Left            =   3960
      TabIndex        =   10
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ALIVE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "K E E P"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   1185
      Left            =   2280
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":1D04
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   975
      Left            =   2280
      TabIndex        =   9
      Top             =   1410
      Width           =   4215
   End
   Begin VB.Label Switch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      ToolTipText     =   " Expand "
      Top             =   2160
      Width           =   165
   End
   Begin VB.Label lblRefresh 
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh (Seconds):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   555
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label lblSecs 
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   795
   End
   Begin VB.Label lblClicks 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Clicks:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   585
   End
   Begin VB.Label Clicks 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Status 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   600
      Width           =   1950
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   120
      Top             =   1560
      Width           =   1950
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayOpen 
         Caption         =   "Open Keep Alive"
      End
      Begin VB.Menu mnuTrayEnable 
         Caption         =   "Enable"
      End
      Begin VB.Menu mnuTraySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Option Explicit


Private Sub chkKeepAwake_Click()
If chkKeepAwake.Value <> 0 Then
    frmMain.tmrUpdate.Enabled = True
    Clicks.Caption = 0
    Status.Caption = 0
Else
    frmMain.tmrUpdate.Enabled = False
End If

End Sub

Private Sub chkTray_Click()
SaveSetting App.EXEName, "Options", "Tray", chkTray.Value
End Sub

Private Sub Form_Load()
chkTray.Value = GetSetting(App.EXEName, "Options", "Tray", 1)
scrSecs.Value = GetSetting(App.EXEName, "Options", "Secs", 1)

' Set timer interval to default of 15 seconds
If scrSecs.Value = 0 Then scrSecs.Value = 15
If scrSecs.Value = 1 Then scrSecs.Value = 15

With nid
  .cbSize = Len(nid)
  .hwnd = Me.hwnd
  .uId = vbNull
  .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
  .uCallBackMessage = WM_MOUSEMOVE
  .hIcon = frmMain.Icon
  .szTip = "Keep Alive" & vbNullChar
End With

Shell_NotifyIcon NIM_ADD, nid

If chkTray.Value = "1" Then Me.Hide

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'This procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim msg As Long

'The value of X will vary depending upon the scalemode setting
If Me.ScaleMode = vbPixels Then
    msg = X
Else
    msg = X / Screen.TwipsPerPixelX
End If

Select Case msg
    Case WM_LBUTTONUP '514 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Case WM_LBUTTONDBLCLK '515 restore form window
    Me.WindowState = vbNormal
    Result = SetForegroundWindow(Me.hwnd)
    Me.Show
    Case WM_RBUTTONUP '517 display popup menu
    Result = SetForegroundWindow(Me.hwnd)
    'Make sure that your first menu item is named "mnu_1",
    'otherwise you will get an error below
    Me.PopupMenu Me.mnuTray
End Select

End Sub

Private Sub Form_Resize()

'This is necessary to assure that the minimized window is hidden
If Me.WindowState = vbMinimized Then Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSetting App.EXEName, "Options", "Secs", scrSecs.Value

'This removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid

Dim intCount As Integer
While Forms.Count > 1
'Find first form besides "me" to unload
intCount = 0

While Forms(intCount).Caption = Me.Caption
intCount = intCount + 1
Wend

Unload Forms(intCount)
Wend

Unload Me
End

End Sub

Private Sub Label6_Click()
Call ShellExecute(hwnd, "Open", "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=7882", "", App.Path, 1)
End Sub

Private Sub Label8_Click()
Me.Hide
End Sub

Private Sub Label8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.Top = Label8.Top + 1
End Sub

Private Sub Label8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.Top = Label8.Top - 1
End Sub

Private Sub mnuTrayEnable_Click()
Me.WindowState = vbNormal
SetForegroundWindow (Me.hwnd)
Me.Show
chkKeepAwake.Value = "1"
End Sub

Private Sub mnuTrayExit_Click()
SaveSetting App.EXEName, "Options", "Secs", scrSecs.Value

'This removes the icon from the system tray
Shell_NotifyIcon NIM_DELETE, nid

Dim intCount As Integer
While Forms.Count > 1
'Find first form besides "me" to unload
intCount = 0

While Forms(intCount).Caption = Me.Caption
intCount = intCount + 1
Wend

Unload Forms(intCount)
Wend

Unload Me
End
End Sub

Private Sub mnuTrayOpen_Click()
Me.WindowState = vbNormal
SetForegroundWindow (Me.hwnd)
Me.Show
End Sub

Private Sub scrSecs_Change()
' When scroll bar changed change timer interval in realtime
txtSecs.Text = scrSecs.Value
tmrUpdate.Interval = scrSecs.Value * 1000
End Sub

Private Sub Switch_Click()
If Switch.Caption = "4" Then
frmMain.Width = "6690"
Switch.Caption = "3"
Switch.Left = "430"
Switch.ToolTipText = ""
Else
frmMain.Width = "2280"
Switch.Caption = "4"
Switch.Left = "136"
Switch.ToolTipText = "Expand"
End If
End Sub

Private Sub Switch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Switch.Caption = "4" Then
Switch.Left = Switch.Left + 1
Else
Switch.Left = Switch.Left - 1
End If
End Sub

Private Sub Switch_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Switch.Caption = "4" Then
Switch.Left = Switch.Left - 1
Else
Switch.Left = Switch.Left + 1
End If
End Sub

Private Sub tmrUpdate_Timer()

' declare for-next variable
Dim xscan As Long

' Put mouse cursor in a 'ready' position
MouseMove (frmMain.Left / Screen.TwipsPerPixelX), (frmMain.Top / Screen.TwipsPerPixelY) + ((frmMain.Height / Screen.TwipsPerPixelY) / 2)

' Begin moving the mouse to 50 random points on the form
For xscan = 1 To 50
    Mouse.MouseMove (frmMain.Left / Screen.TwipsPerPixelX) + (Rnd * (frmMain.Width / Screen.TwipsPerPixelX)), (frmMain.Top / Screen.TwipsPerPixelY) + (Rnd * (frmMain.Height / Screen.TwipsPerPixelY))
Next

' Click on the form to prevent AllAdvantage time out
MouseMove ((frmMain.Left / Screen.TwipsPerPixelX) + (frmMain.Width / Screen.TwipsPerPixelX) / 2), _
            ((frmMain.Top / Screen.TwipsPerPixelY) + (frmMain.Height / Screen.TwipsPerPixelY) / 2)
MouseFullClick (btcMiddle)
Clicks.Caption = Clicks.Caption + 1
Mouse.MouseFullClick (btcMiddle)

' Increment Seconds Counter
Status.Caption = Status.Caption + Val(txtSecs.Text)
End Sub
