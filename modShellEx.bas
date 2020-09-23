Attribute VB_Name = "modShellEx"
Option Explicit

'# function to launch shortcuts
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'# consts for shell execute errors
Private Const ERROR_FILE_NOT_FOUND   As Long = 2&
Private Const ERROR_PATH_NOT_FOUND   As Long = 3&
Private Const ERROR_BAD_FORMAT       As Long = 11&
Private Const SE_ERR_ACCESSDENIED    As Long = 5
Private Const SE_ERR_ASSOCINCOMPLETE As Long = 27
Private Const SE_ERR_DDEBUSY         As Long = 30
Private Const SE_ERR_DDEFAIL         As Long = 29
Private Const SE_ERR_DDETIMEOUT      As Long = 28
Private Const SE_ERR_DLLNOTFOUND     As Long = 32
Private Const SE_ERR_FNF             As Long = 2
Private Const SE_ERR_NOASSOC         As Long = 31
Private Const SE_ERR_OOM             As Long = 8
Private Const SE_ERR_PNF             As Long = 3
Private Const SE_ERR_SHARE           As Long = 26


Public Sub LaunchLink(Frm As Form, Lnk As String)
Dim lngRet As Long

    PlayThisSound STSND_LAUNCH                  '# play the corresponding sound
    
    '# launch the shortcut
    lngRet = ShellExecute(Frm.hwnd, vbNullString, Lnk, vbNullString, vbNullString, vbNormalFocus)                    '# shell shortcut
    
    '# anything over 32 is a success
    '# I could only find 14 error codes explained - everything else is unknown
    If UseTips Then
        Select Case lngRet
            Case 0:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The operating system is out of memory or resources.", ICO_EXCLAMATION, 2000
            Case ERROR_FILE_NOT_FOUND:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The specified file was not found.", ICO_EXCLAMATION, 2000
            Case ERROR_PATH_NOT_FOUND:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The specified path was not found.", ICO_EXCLAMATION, 2000
            Case ERROR_BAD_FORMAT:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The .exe file is invalid (non-Win32® .exe or error in .exe image).", ICO_EXCLAMATION, 2000
            Case SE_ERR_ACCESSDENIED:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The operating system denied access to the specified file.", ICO_EXCLAMATION, 2000
            Case SE_ERR_ASSOCINCOMPLETE:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The file name association is incomplete or invalid.", ICO_EXCLAMATION, 2000
            Case SE_ERR_DDEBUSY:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The DDE transaction could not be completed because other DDE transactions were being processed.", ICO_EXCLAMATION, 2000
            Case SE_ERR_DDEFAIL:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The DDE transaction failed.", ICO_EXCLAMATION, 2000
            Case SE_ERR_DDETIMEOUT:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The DDE transaction could not be completed because the request timed out.", ICO_EXCLAMATION, 2000
            Case SE_ERR_DLLNOTFOUND:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The specified dynamic-link library was not found.", ICO_EXCLAMATION, 2000
            Case SE_ERR_FNF:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The specified file was not found.", ICO_EXCLAMATION, 2000
            Case SE_ERR_NOASSOC:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "There is no application associated with the given file name extension." & vbCrLf & "This error will also be returned if you attempt to print a file that is not printable.", ICO_EXCLAMATION, 2000
            Case SE_ERR_OOM:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "There was not enough memory to complete the operation.", ICO_EXCLAMATION, 2000
            Case SE_ERR_PNF:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "The specified path was not found.", ICO_EXCLAMATION, 2000
            Case SE_ERR_SHARE:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "A sharing violation occurred.", ICO_EXCLAMATION, 2000
            Case Is > 32:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " launched.", ICO_INFO, 2000
            Case Else:
                ModifyTray Frm.hwnd, Frm.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Shortcut : " & Left(Lnk, Len(Lnk) - 5) & " failed to launch." & vbCrLf & "An unknown error has occurred.", ICO_EXCLAMATION, 2000
        End Select
    End If
    
End Sub
