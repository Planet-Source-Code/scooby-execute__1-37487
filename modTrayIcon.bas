Attribute VB_Name = "modTrayIcon"
Option Explicit

'# api functions for controlling systray icon
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'# consts for systray icon manipulation
Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD            As Long = &H0
Private Const NIM_MODIFY         As Long = &H1
Private Const NIM_DELETE         As Long = &H2
Private Const NIM_SETFOCUS       As Long = &H3
Private Const NIM_SETVERSION     As Long = &H4
Private Const NIF_MESSAGE        As Long = &H1
Private Const NIF_ICON           As Long = &H2
Private Const NIF_TIP            As Long = &H4
Private Const NIF_STATE          As Long = &H8
Private Const NIF_INFO           As Long = &H10
Private Const NIS_HIDDEN         As Long = &H1
Private Const NIS_SHAREDICON     As Long = &H2
Private Const NIIF_NONE          As Long = &H0
Private Const NIIF_WARNING       As Long = &H2
Private Const NIIF_ERROR         As Long = &H3
Private Const NIIF_INFO          As Long = &H1
Private Const NIIF_GUID          As Long = &H4

'# Enum for the ballon icons
Enum BalloonIcons
    ICO_NONE = 0&
    ICO_INFO = 1&
    ICO_EXCLAMATION = 2&
    ICO_WARNING = 3&
End Enum

'# type used in the systray icon procedures
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type
Private nid As NOTIFYICONDATA

Public Sub AddToTray(TrayHandle As Long, TrayIconImage As Long, TrayText As String, Optional BalloonTitle As String, Optional BalloonText As String, Optional BalloonIcon As BalloonIcons, Optional TrayTimeOut As Long = 3000)
        
    On Error GoTo Err_Handler
    
    With nid
        .cbSize = Len(nid)                      '# set the size of the data structure
        If TrayHandle <> 0 Then
            .hwnd = TrayHandle                  '# get the handle of the item to receive callbacks
        End If
        .uId = vbNull
        If BalloonTitle <> "" Then              '# if we are using balloon type tooltips
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP    '# set appropriate flags
        Else
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        End If
        .uCallBackMessage = WM_MOUSEMOVE        '# the message sent to the callback function
        If TrayIconImage <> 0 Then              '# if we have added an icon to the tip
            .hIcon = TrayIconImage
        End If
        If TrayText <> "" Then                  '# if we have added text to the tip
            .szTip = TrayText & vbNullChar
        End If
        .dwState = 0
        .dwStateMask = 0
        If BalloonTitle <> "" Or BalloonText <> "" Or BalloonIcon <> 0 Then
            If BalloonText <> "" Then           '# text in the balloon tip
                .szInfo = BalloonText & Chr(0)  '# add null char
            End If
            If BalloonTitle <> "" Then          '# title for the balloon tip
                .szInfoTitle = BalloonTitle & Chr(0)    '# add null char
            End If
            If BalloonIcon <> 0 Then            '# icon used in balloon tips
                .dwInfoFlags = BalloonIcon
            End If
            If TrayTimeOut > 0 Then             '# amount of time to display tip
                .uTimeout = TrayTimeOut
            End If
        End If
    End With
    Shell_NotifyIcon NIM_ADD, nid               '# add icon to tray

    Exit Sub
    
Err_Handler:
    ErrPosition = "AddToTray()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub

Public Sub RemoveFromTray()

    On Error GoTo Err_Handler
    
    Shell_NotifyIcon NIM_DELETE, nid            '# remove icon from tray
    
    Exit Sub
    
Err_Handler:
    ErrPosition = "RemoveFromTray()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Public Sub ModifyTray(Optional TrayHandle As Long, Optional TrayIconImage As Long, Optional TrayText As String, Optional BalloonTitle As String, Optional BalloonText As String, Optional BalloonIcon As Long, Optional TrayTimeOut As Long)

    On Error GoTo Err_Handler
    
    With nid
        .cbSize = Len(nid)                              '# set size of data structure
        If TrayHandle <> 0 Then                         '# handle of item to receive callback messages
            .hwnd = TrayHandle
        End If
        .uId = vbNull
        If BalloonTitle <> "" Then                      '# are we using new style balloon tips
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP    '# set appropriate flags
        Else
            .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
        End If
        .uCallBackMessage = WM_MOUSEMOVE                '# message to pass to callback function
        .hIcon = TrayIconImage                          '# systray icon
        If TrayText <> "" Then                          '# systray tip text
            .szTip = TrayText & vbNullChar              '# add the null char
        End If
        .dwState = 0
        .dwStateMask = 0
        If BalloonTitle <> "" Or BalloonText <> "" Or BalloonIcon <> 0 Then '# are we using new style balloon tips
            If BalloonText <> "" Then                   '# set balloon text
                .szInfo = BalloonText & Chr(0)          '# add null char
            End If
            If BalloonTitle <> "" Then                  '# set balloon title
                .szInfoTitle = BalloonTitle & Chr(0)    '# add null char
            End If
            If BalloonIcon <> 0 Then                    '# set balloon icon
                .dwInfoFlags = BalloonIcon
            End If
            If TrayTimeOut > 0 Then                     '# set timeout
                .uTimeout = TrayTimeOut
            End If
        End If
    End With
    Shell_NotifyIcon NIM_MODIFY, nid                    '# modify current systray icon
    
    Exit Sub
    
Err_Handler:
    ErrPosition = "ModifyTray()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub
