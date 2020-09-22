VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSpy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SPY WIN"
   ClientHeight    =   4815
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   4905
   Icon            =   "frmSpy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   555
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   48
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   47
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   195
      Left            =   2400
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Show MyContrlo"
      Height          =   375
      Left            =   240
      TabIndex        =   46
      Top             =   4320
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog Cdbox 
      Left            =   3480
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "MY Control"
      Height          =   4335
      Index           =   0
      Left            =   4920
      TabIndex        =   29
      Top             =   240
      Width           =   2175
      Begin VB.Frame Frame6 
         Caption         =   "Wall Paper"
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   3480
         Width           =   1695
         Begin VB.CommandButton Command3 
            Caption         =   "Set Pic"
            Height          =   255
            Left            =   360
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Start"
         Height          =   615
         Left            =   240
         TabIndex        =   41
         Top             =   2760
         Width           =   1695
         Begin VB.CommandButton Command4 
            Caption         =   "Enable"
            Height          =   255
            Left            =   840
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Disable"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Time"
         Height          =   615
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   1695
         Begin VB.CommandButton Command7 
            Caption         =   "Show"
            Height          =   255
            Left            =   840
            TabIndex        =   38
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Hide"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "CD-Rom"
         Height          =   615
         Left            =   240
         TabIndex        =   33
         Top             =   2040
         Width           =   1695
         Begin VB.CommandButton Command6 
            Caption         =   "Close"
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Open"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "TaskBar"
         Height          =   855
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton Command10 
            Caption         =   "Enable"
            Height          =   255
            Left            =   840
            TabIndex        =   40
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Disable"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Show"
            Height          =   255
            Left            =   840
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Hide"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Lock Mode"
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   27
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   25
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   23
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   195
      Left            =   2400
      TabIndex        =   20
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   3000
      Top             =   4680
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drag Me"
      Enabled         =   0   'False
      Height          =   915
      Left            =   3240
      TabIndex        =   16
      Top             =   480
      Width           =   825
      Begin VB.PictureBox picDrag 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   240
         MouseIcon       =   "frmSpy.frx":0442
         Picture         =   "frmSpy.frx":074C
         ScaleHeight     =   360
         ScaleWidth      =   315
         TabIndex        =   17
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000006&
      Height          =   555
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   0
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Mouse pos. Y:"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label13 
      Caption         =   "Mouse pos. X:"
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Window state:"
      Height          =   195
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Window height:"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Window width:"
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Window Text:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Window Class Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Window Handle:"
      Height          =   180
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Window Style:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Window ID Number:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   1440
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Parent Window Handle:"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Parent Window Text :"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Parent Window Class Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   2025
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Module:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   570
   End
   Begin VB.Menu mnuDoStuff 
      Caption         =   "Contrlo Win"
      Begin VB.Menu mnuEnableDisable 
         Caption         =   "Enable/ Disable"
         Begin VB.Menu mnuEnable 
            Caption         =   "Enable"
         End
         Begin VB.Menu mnuDisable 
            Caption         =   "Disable"
         End
      End
      Begin VB.Menu mnuSet 
         Caption         =   "Set"
         Begin VB.Menu mnuOnTopVals 
            Caption         =   "On top values"
            Begin VB.Menu mnuWinOnTop 
               Caption         =   "Window on top"
            End
            Begin VB.Menu mnuWinNotOnTop 
               Caption         =   "Window not on top"
            End
         End
         Begin VB.Menu mnuZOrderPos 
            Caption         =   "Z-Order position"
            Begin VB.Menu mnuZTop 
               Caption         =   "Top of Z-order"
            End
            Begin VB.Menu mnuZbottom 
               Caption         =   "Bottom of Z-order"
            End
         End
         Begin VB.Menu mnuChangeCText 
            Caption         =   "Change control text"
         End
         Begin VB.Menu mnuChangeText 
            Caption         =   "Change window text"
         End
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send Message to window"
         Begin VB.Menu mnuCreate 
            Caption         =   "Create"
         End
         Begin VB.Menu mnuDestroy 
            Caption         =   "Destroy"
         End
         Begin VB.Menu mnuClose 
            Caption         =   "Close"
         End
         Begin VB.Menu mnuRefresh 
            Caption         =   "Refresh"
         End
         Begin VB.Menu mnuClick 
            Caption         =   "Click"
            Begin VB.Menu mnuLeftClick 
               Caption         =   "Left click"
            End
            Begin VB.Menu mnuLeftDblClick 
               Caption         =   "Left click (double)"
            End
            Begin VB.Menu mnuRightClick 
               Caption         =   "Right click"
            End
            Begin VB.Menu mnuRightDblClick 
               Caption         =   "Right click (double)"
            End
         End
      End
      Begin VB.Menu mnuShowWindow 
         Caption         =   "Show window"
         Begin VB.Menu mnuHideWin 
            Caption         =   "Hide window"
         End
         Begin VB.Menu mnuShowWin 
            Caption         =   "Show window"
         End
         Begin VB.Menu mnuMinimize 
            Caption         =   "Minimize"
         End
         Begin VB.Menu mnuMaximize 
            Caption         =   "Maximize"
         End
         Begin VB.Menu mnuRestore 
            Caption         =   "Restore"
         End
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuAboutPro 
         Caption         =   "AboutPro"
      End
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Long
Dim Flag As Boolean


Private Sub Check1_Click()
    
    If Frame1.Enabled = False Then
    Frame1.Enabled = True
    Timer1.Enabled = False
    Exit Sub
    End If
    Frame1.Enabled = False
    Timer1.Enabled = True
    
End Sub

Private Sub Command1_Click()
Call ShowTaskBar
End Sub

Private Sub Command10_Click()
Call EnableTaskBar
End Sub

Private Sub Command11_Click()
DisableStart (True)
End Sub

Private Sub Command12_Click()
If Flag = False Then
Command12.Caption = "Hide MyContrlo"
Flag = True
Me.Width = 7350
ElseIf Flag = True Then
Command12.Caption = "Show MyContrlo"
Flag = False
Me.Width = 5000
End If
End Sub

Private Sub Command2_Click()
Call HideTaskBar
End Sub

Private Sub Command3_Click()
    Cdbox.DialogTitle = "Choose a bitmap"
    Cdbox.Filter = "Windows Bitmaps (*.BMP)|*.bmp"
    Cdbox.ShowOpen
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0, Cdbox.FileName, SPIF_UPDATEINIFILE
End Sub

Private Sub Command4_Click()
DisableStart (False)
End Sub

Private Sub Command5_Click()
Call OpenCDROM
End Sub

Private Sub Command6_Click()
Call CloseCDROM
End Sub

Private Sub Command7_Click()
Call ShowTime
End Sub

Private Sub Command8_Click()
Call HideTime
End Sub

Private Sub Command9_Click()
Call DisableTaskBar
End Sub

Private Sub Form_Activate()
SetWindowPos Me.hWnd, -1, 0, 0, 0, 0, &H10 Or &H40 Or &H2 Or &H1
End Sub

Private Sub Form_Load()
Me.Icon = LoadResPicture(101, vbResIcon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
NotOnTop Me
MsgBox "don't forget to vote plase"
End Sub

Private Sub mnuAboutPro_Click()
 NotOnTop Me
 frmAbout.Show
OnTop frmAbout
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDrag.MousePointer = 99
    Me.MousePointer = 99
    picDrag.Picture = Me.Picture
    InformationNow = True
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InformationNow = True Then
    Call spy
    End If
End Sub

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDrag.MousePointer = 0
    Me.MousePointer = 0
    picDrag.Picture = LoadResPicture(101, vbResIcon)
    InformationNow = False
    Call spy
End Sub
Private Sub mnuChangeCText_Click()
    Dim Input_ As String
    NotOnTop Me
    Input_ = InputBox("Change control text to:", "Change window text", MainText)
    OnTop Me
    Call SetWindowText(Win, Input_)
    Call SendMessage(Win, WM_SETTEXT, ByVal CLng(0), ByVal Input_)
End Sub

Private Sub mnuChangeText_Click()
    Dim Input_ As String
    NotOnTop Me
    Input_ = InputBox("Change text to:", "Change window text", MainText)
    OnTop Me
    Call SetWindowText(Win, Input_)
    Call SendMessage(Win, WM_PAINT, 0&, 0&)
End Sub

Private Sub mnuClose_Click()
    Call SendMessage(Win, WM_CLOSE, 0, 0)
End Sub

Private Sub mnuCreate_Click()
    Call SendMessage(Win, WM_CREATE, 0&, 0&)
End Sub

Private Sub mnuDestroy_Click()
    Call SendMessage(Win, WM_DESTROY, 0&, 0&)
End Sub

Private Sub mnuDisable_Click()
    Call EnableWindow(Win, EW_DISABLE)
End Sub

Private Sub mnuEnable_Click()
    Call EnableWindow(Win, EW_Enable)
End Sub

Private Sub mnuHideWin_Click()
    Call ShowWindow(Win, SW_HIDE)
End Sub

Private Sub mnuLeftClick_Click()
    Call PostMessage(Win, WM_LBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
    Call PostMessage(Win, WM_LBUTTONUP, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuLeftDblClick_Click()
   Call PostMessage(Win, WM_LBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuMaximize_Click()
    Call ShowWindow(Win, SW_MAXIMIZE)
End Sub

Private Sub mnuMinimize_Click()
    Call ShowWindow(Win, SW_MINIMIZE)
End Sub

Private Sub mnuRefresh_Click()
    Call SendMessage(Win, WM_PAINT, 0, 0)
End Sub

Private Sub mnuRestore_Click()
    Call ShowWindow(Win, SW_RESTORE)
End Sub

Private Sub mnuRightClick_Click()
    Call PostMessage(Win, WM_RBUTTONDOWN, ByVal CLng(0), ByVal CLng(0))
    Call PostMessage(Win, WM_RBUTTONUP, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuRightDblClick_Click()
    Call PostMessage(Win, WM_RBUTTONDBLCLK, ByVal CLng(0), ByVal CLng(0))
End Sub

Private Sub mnuShowWin_Click()
    Call ShowWindow(Win, SW_SHOW)
End Sub

Private Sub mnuWinNotOnTop_Click()
    Call SetWindowPos(Win, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuWinOnTop_Click()
    Call SetWindowPos(Win, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuZbottom_Click()
  Call SetWindowPos(Win, HWND_BOTTOM, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Private Sub mnuZTop_Click()
    Call SetWindowPos(Win, HWND_TOP, 0, 0, 0, 0, SWP_FLAGS)
End Sub

Sub spy()

WindowSPY Text1, Text2, Text3, Text4, Text5, Text6, Text7, Text8, Text9
Win = Val(Text1.Text)
    
    Call GetWindowRect(Win, RectMain)
    Text10.Text = RectMain.Right - RectMain.Left
    Text11.Text = RectMain.Bottom - RectMain.Top
    
    If (Not IsIconic(Win)) And (Not IsZoomed(Win)) Then Text12.Text = "General"
    If IsIconic(Win) Then Text12.Text = "Minimized"
    If IsZoomed(Win) Then Text12.Text = "Maximized"
    
    Call GetCursorPos(CurPos)
    Text13.Text = CurPos.X
    Text14.Text = CurPos.Y
    
End Sub



Private Sub Timer1_Timer()
Call spy
End Sub

Sub WindowSPY(WinHdl As TextBox, WinClass As TextBox, _
WinTxt As TextBox, WinStyle As TextBox, _
WinIDNum As TextBox, WinPHandle As TextBox, _
WinPText As TextBox, WinPClass As TextBox, _
WinModule As TextBox)

    'Call This In A Timer
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
    Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
    Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
    Dim hInstance As Long, sParentWindowText As String * 100
    Dim sModuleFileName As String * 100, r As Long
    
    Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.Y
    
    hWndOver = WindowFromPoint(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Text = Str(hWndOver)
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxt.Text = Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.Text = Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Text = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)


        If hWndParent <> 0 Then
            wID = GetWindowWord(hWndOver, GWW_ID)
            WinIDNum.Text = "Window ID Number: " & wID
            WinPHandle.Text = hWndParent
            r = GetWindowText(hWndParent, sParentWindowText, 100)
            WinPText.Text = Left(sParentWindowText, r)
            r = GetClassName(hWndParent, sParentClassName, 100)
            WinPClass.Text = Left(sParentClassName, r)
        Else
            WinIDNum.Text = "N/A"
            WinPHandle.Text = "N/A"
            WinPText.Text = "N/A"
            WinPClass.Text = "N/A"
        End If

        hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
        r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Text = Left(sModuleFileName, r)
    End If

End Sub


