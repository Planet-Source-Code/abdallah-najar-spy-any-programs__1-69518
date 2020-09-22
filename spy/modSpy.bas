Attribute VB_Name = "Module1"

Public Win As Long
Public CurPos As POINTAPI
Public RectMain As Rect

Public Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Const EW_Enable = 1
Public Const EW_DISABLE = 0


Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_STYLE = (-16)

Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As Rect) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long

Public Declare Function IsIconic Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32.dll" (ByVal hWnd As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_CREATE = 1
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_DESTROY = 2
Public Const WM_MOVE = 3
Public Const WM_SIZE = 5
Public Const WM_PAINT = &HF
Public Const WM_DRAGFORM = &HA1

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_UPDATEINIFILE = &H1

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const SW_RESTORE = 9

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_BOTTOM = 1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_DRAWFRAME = &H20
Public Const SWP_FLAGS = SWP_NOSIZE Or SWP_NOMOVE

Public Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Public Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Sub HideTaskBar()
    Dim Taskbar As Long
        Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Taskbar&, SW_HIDE)
End Sub

Public Sub ShowTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Taskbar&, SW_SHOW)
End Sub

Public Sub HideTime()
    Dim ParentOfAnnoyance As Long, Child As Long, Annoyance As Long
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "TrayNotifyWnd", vbNullString)
    Annoyance = FindWindowEx(Child&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(Annoyance&, SW_HIDE)
End Sub

Public Sub ShowTime()
    Dim ParentOfAnnoyance As Long, Child As Long, Annoyance As Long
    ParentOfAnnoyance = FindWindow("Shell_TrayWnd", vbNullString)
    Child = FindWindowEx(ParentOfAnnoyance&, 0&, "TrayNotifyWnd", vbNullString)
    Annoyance = FindWindowEx(Child&, 0&, "TrayClockWClass", vbNullString)
    Call ShowWindow(Annoyance&, SW_SHOW)
End Sub

Public Function DisableTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call EnableWindow(Taskbar&, EW_DISABLE)
End Function

Public Function EnableTaskBar()
    Dim Taskbar As Long
    Taskbar = FindWindow("Shell_TrayWnd", vbNullString)
    Call EnableWindow(Taskbar&, EW_Enable)
End Function

Sub OpenCDROM()
    Dim CD_ROM
    CD_ROM = mciSendString("set CDAudio door open", RetString, 0&, 0&)
End Sub

Sub CloseCDROM()
    Dim CD_ROM
    CD_ROM = mciSendString("set CDAudio door closed", 0&, 0&, 0&)
End Sub

Public Sub OnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub NotOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Public Sub DisableStart(Disable As Boolean)
    Dim Taskbar As Long, StartButton As Long
    Taskbar& = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton& = FindWindowEx(Taskbar&, 0&, "Button", vbNullString)
    If Disable = True Then
        Call EnableWindow(StartButton&, EW_DISABLE)
    ElseIf Disable = False Then
        Call EnableWindow(StartButton&, EW_Enable)
    End If
End Sub
