Attribute VB_Name = "modTaskBarForm"
'Download by http://www.NewXing.com
Option Explicit

'---API Function and Sub declarations
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'---Type declarations
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'---Constant declarations
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

'---Variable declarations
Private rectLastTray As RECT
Private rectLastRebar As RECT
Private rectLastNotify As RECT
Private hwndForm As Long
Private lngTimer As Long
Private intWidth As Integer
Private intHeight As Integer
Private blnGrow As Boolean



Public Sub AttachForm(MyForm As Form, Optional intForceWidth As Integer = 0, Optional intForceHeight As Integer = 0, Optional blnGrowWithTray As Boolean = False)
    '---Set the variable to the handle of the form being passed in
    hwndForm = MyForm.hWnd
    
    '---Set the width of the encapsulation
    If intForceWidth <> 0 Then
        intWidth = intForceWidth
    Else
        intWidth = MyForm.Width
    End If
    
    '---Set the height of the encapsulation
    If intForceHeight <> 0 Then
        intHeight = intForceHeight
    Else
        intHeight = MyForm.Height
    End If
    
    '---Set grow with tray option
    blnGrow = blnGrowWithTray
    
    '---Set the parent window to Shell_TrayWnd
    SetParent hwndForm, GetTrayHandle
    
    '---Start the timer process
    lngTimer = SetTimer(hwndForm, 0, 50, AddressOf MainLoop)

End Sub

Public Sub DetachForm()
    Dim rectTray As RECT
    Dim rectTrayClient As RECT
    Dim rectRebar As RECT
    Dim rectNotify As RECT
    'variables for window placement & size
    'these are only used because it makes it easier to see the calculations
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim h As Long
    
    '---Get the current RECT structures of the main taskbar windows (Shell_TrayWnd, ReBarWindow32, TrayNotifyWnd)
    GetWindowRect GetTrayHandle, rectTray
    GetClientRect GetTrayHandle, rectTrayClient
    GetWindowRect GetRebarHandle, rectRebar
    GetWindowRect GetNotifyHandle, rectNotify
    
    '---Sets the main forms parent to be the desktop
    SetParent hwndForm, vbNull
    '---Kill the system timer
    KillTimer hwndForm, lngTimer

    '---Set the ReBarWindow32 back to normal
        If (rectTray.Right - rectTray.Left) = (Screen.Width / Screen.TwipsPerPixelX) Then   'Horizontal
            Debug.Print "Taskbar Orientation: Horizontal"
            '---Resize the ReBarWindow32 and refresh its RECT structure
            X = rectRebar.Left - rectTray.Left              'original starting position
            Y = rectTrayClient.Top                          'always at the top
            w = rectNotify.Left - rectRebar.Left            'align with notify tray
            h = rectRebar.Bottom - rectRebar.Top            'original height
            MoveWindow GetRebarHandle, X, Y, w, h, 1
            GetWindowRect GetRebarHandle, rectRebar
        ElseIf (rectTray.Bottom - rectTray.Top) = (Screen.Height / Screen.TwipsPerPixelY) Then
            Debug.Print "Taskbar Orientation: Vertical"
            '---Resize the ReBarWindow32 and refresh its RECT structure
            X = rectTrayClient.Left                         'always at left
            Y = rectRebar.Top - rectTray.Top                'original starting y
            h = rectNotify.Top - rectRebar.Top              'align with notify tray
            w = rectRebar.Right - rectRebar.Left            'original width
            MoveWindow GetRebarHandle, X, Y, w, h, 1
            GetWindowRect GetRebarHandle, rectRebar
        End If
End Sub


Sub MainLoop()
    Dim rectTray As RECT
    Dim rectTrayClient As RECT
    Dim rectRebar As RECT
    Dim rectNotify As RECT
    'variables for window placement & size
    'these are only used because it makes it easier to see the calculations
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim h As Long
    
    On Error Resume Next
    
    DoEvents

    '---Get the current RECT structures of the main taskbar windows (Shell_TrayWnd, ReBarWindow32, TrayNotifyWnd)
    GetWindowRect GetTrayHandle, rectTray
    GetClientRect GetTrayHandle, rectTrayClient
    GetWindowRect GetRebarHandle, rectRebar
    GetWindowRect GetNotifyHandle, rectNotify
    
    'Debug.Print "-------------- " & Now & "---------------------------------------"
    'Debug.Print "Shell_TrayWnd   -->  L:" & rectTray.Left & " R:" & rectTray.Right & " T:" & rectTray.Top & " B:" & rectTray.Bottom
    'Debug.Print "Shell_TrayWnd(C)-->  L:" & rectTrayClient.Left & " R:" & rectTrayClient.Right & " T:" & rectTrayClient.Top & " B:" & rectTrayClient.Bottom
    'Debug.Print "ReBarWindow32   -->  L:" & rectRebar.Left & " R:" & rectRebar.Right & " T:" & rectRebar.Top & " B:" & rectRebar.Bottom
    'Debug.Print "TrayNotifyWnd   -->  L:" & rectNotify.Left & " R:" & rectNotify.Right & " T:" & rectNotify.Top & " B:" & rectNotify.Bottom

    '--- Check to see if any of the RECT structures have changed (Task Bar has been resized, icon added to notification area, new app opened, etc)
    'these comparisons worked during all my testing but you could also check for rectTray.Bottom, etc if taskbar is on top of screen
    If rectTray.Top <> rectLastTray.Top Or rectRebar.Right <> rectLastRebar.Right Or rectNotify.Left <> rectLastNotify.Left Then
        Debug.Print "Task bar window(s) resized...  Recalcualting position."
        
        '---Determine orientation and location of Shell_TrayWnd
        If (rectTray.Right - rectTray.Left) > (rectTray.Bottom - rectTray.Top) Then   'Horizontal
            Debug.Print "Taskbar Orientation: Horizontal"
            
            '---Resize the ReBarWindow32 and refresh its RECT structure
            X = rectRebar.Left - rectTray.Left              'original starting position
            Y = rectTrayClient.Top                          'always at the top
            w = rectNotify.Left - rectRebar.Left - intWidth 'put a buffer between the notify and rebar windows
            h = rectRebar.Bottom - rectRebar.Top            'original height
            MoveWindow GetRebarHandle, X, Y, w, h, 1
            GetWindowRect GetRebarHandle, rectRebar
            
            '---This moves our form into position on the task bar
            X = rectRebar.Right                             'start at right of rebar
            Y = rectTrayClient.Top + 4                      'give a 4 pixel buffer from top of tray client area
            w = intWidth                                    'width as specified
            'because we are horizontal, we need to check to see if the height specified is larger then the client area of the tray.  if so, truncate height.
            If (intHeight > (rectTrayClient.Bottom - rectTrayClient.Top - 6)) Or blnGrow = True Then
                h = rectTrayClient.Bottom - rectTrayClient.Top - 6
            Else
                h = intHeight
            End If
            MoveWindow hwndForm, X, Y, w, h, 1
        
        ElseIf (rectTray.Bottom - rectTray.Top) > (rectTray.Right - rectTray.Left) Then 'Vertical
            Debug.Print "Taskbar Orientation: Vertical"
            
            '---Resize the ReBarWindow32 and refresh its RECT structure
            X = rectTrayClient.Left                         'always at left
            Y = rectRebar.Top - rectTray.Top                'original starting y
            h = rectNotify.Top - rectRebar.Top - intHeight  'specified height
            w = rectRebar.Right - rectRebar.Left            'original width
            MoveWindow GetRebarHandle, X, Y, w, h, 1
            GetWindowRect GetRebarHandle, rectRebar
            
            '---This moves our form into position on the task bar
            X = rectTrayClient.Left + 4
            Y = rectRebar.Bottom
            h = intHeight                       'specified height
            'becasue we are vertical, we need to check if the width specified is larger r\than the client area of the tray, if so, truncate it.
            If (intWidth > (rectTrayClient.Right - rectTrayClient.Left - 6)) Or blnGrow = True Then
                w = rectTrayClient.Right - rectTrayClient.Left - 6
            Else
                w = intWidth
            End If
            MoveWindow hwndForm, X, Y, w, h, 1
        End If
                
        '---Sets the form so it's always on top
        'SetWindowPos frmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    End If
    
    rectLastTray = rectTray
    rectLastRebar = rectRebar
    rectLastNotify = rectNotify
End Sub


Private Function GetTrayHandle() As Long
    '---This function returns the hWnd of the Shell_TrayWnd window (the whole tak bar)
    Dim hWnd_Tray As Long
    
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    GetTrayHandle = hWnd_Tray
End Function

Private Function GetRebarHandle() As Long
    '---This function returns the hWnd of the ReBarWindow32 windo (task bar buttons area, quicklaunch, etc)
    Dim hWnd_Tray As Long
    Dim hWnd_Rebar As Long
    
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    
    If hWnd_Tray <> 0 Then
        hWnd_Rebar = FindWindowEx(hWnd_Tray&, 0, "ReBarWindow32", vbNullString)
    End If
    
    GetRebarHandle = hWnd_Rebar
End Function

Function GetNotifyHandle() As Long
    '---This function simply returns the hWnd of the TrayNotifyWnd window (the notification icon area)
    Dim hWnd_Tray As Long
    Dim hWnd_Notify As Long
    
    hWnd_Tray = FindWindow("Shell_TrayWnd", "")
    
    If hWnd_Tray <> 0 Then
        hWnd_Notify = FindWindowEx(hWnd_Tray&, 0, "TrayNotifyWnd", vbNullString)
    End If
    
    GetNotifyHandle = hWnd_Notify
End Function


