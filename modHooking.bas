Attribute VB_Name = "modHooking"
Option Explicit

'# api finctions
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'# const for hooking
Private Const GWL_WNDPROC = -4

'# windows messages
Public Const WM_MOUSEMOVE       As Long = &H200
Public Const WM_LBUTTONDOWN     As Long = &H201
Public Const WM_LBUTTONUP       As Long = &H202
Public Const WM_LBUTTONDBLCLK   As Long = &H203
Public Const WM_RBUTTONDOWN     As Long = &H204
Public Const WM_RBUTTONUP       As Long = &H205
Public Const WM_RBUTTONDBLCLK   As Long = &H206
Public Const WM_DRAWITEM        As Long = &H2B
Public Const WM_MEASUREITEM     As Long = &H2C

'# used in hooking
Global IsHooked As Boolean
Global gHW As Long

Global lpPrevWndProc As Long
Global lpNewWndProc As Long

Public Sub Hook()
'# hooking means intercepting the Windows messaging system to see what's going on
    
    On Error GoTo Err_Handler
    
    lpNewWndProc = ProcAddr(AddressOf WndProc)              '# set the memory address of the new Windows procedure
    If Not IsHooked Then                                    '# never hook when already hooked as you won't be able to unhook
        '# setting the new default procedure returns the memory address of the old one.  Save it for unhooking
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, lpNewWndProc)
        IsHooked = True                                     '# set the flag to say we are hooked
    End If

    Exit Sub
    
Err_Handler:
    ErrPosition = "Hook()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Private Function ProcAddr(ByVal Address As Long)
'# returns the memory address of a procedure
  
    ProcAddr = Address

End Function

Public Sub Unhook()
'# returns control of windows messages to the default procedure
Dim temp As Long

    On Error GoTo Err_Handler
    
    temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)   '# set the default Windows procedure to the original
    If temp <> 0 Then                                       '# if it worked set flag
        IsHooked = False
    End If

    Exit Sub

Err_Handler:
    ErrPosition = "UnHook()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'# generic replacement window procedure
    
    On Error GoTo Err_Handler
    
    Select Case uMsg                                        '# trap the messages you want to see
        Case WM_MEASUREITEM:
            MeasureMenu lParam
        Case WM_DRAWITEM:
            DrawMenu frmMain, hMenu, lParam, SBTEXT, NumSeps, FirstSep
        Case Else:
    End Select
    
    '# usually a good idea to always let Windows actually process all messages
    '# when you have finished with them
TidyUp:
    WndProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, lParam)
    Exit Function

Err_Handler:
    ErrPosition = "WndProc()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    Resume TidyUp
    
End Function
