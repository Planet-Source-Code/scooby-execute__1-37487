Attribute VB_Name = "modSystemMenu"
Option Explicit

'# api functions to access/modify the system menu
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

'# consts used to access menu items
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Public Sub DisableX(Frm As Form)
Dim hSysMenu As Long
Dim lngCnt As Long
    
    On Error GoTo Err_Handler
    
    '# Get handle to the form's system menu
    '# (Restore, Maximize, Move, close etc.)
    hSysMenu = GetSystemMenu(Frm.hwnd, False)

    If hSysMenu Then
        '# Get System menu's menu count
        lngCnt = GetMenuItemCount(hSysMenu)
        If lngCnt Then
            '# menu count starts at 0
            '# close is last
            RemoveMenu hSysMenu, lngCnt - 1, MF_BYPOSITION Or MF_REMOVE
            '# then a seperator
            RemoveMenu hSysMenu, lngCnt - 2, MF_BYPOSITION Or MF_REMOVE
            DrawMenuBar Frm.hwnd
            '# refresh
            Frm.Refresh
        End If
    End If

    Exit Sub
    
Err_Handler:
    ErrPosition = "DisableX()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub
