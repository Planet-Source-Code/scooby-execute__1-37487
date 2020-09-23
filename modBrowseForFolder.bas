Attribute VB_Name = "modBrowseForFolder"
Option Explicit

'api functions for controlling the browse For folder dialog
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfoType) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function CoInitialize Lib "OLE32.DLL" (ByVal pvReserved As Long) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal pv As Long)
Private Declare Sub CoUninitialize Lib "OLE32.DLL" ()

'# Enum for the Flags of the BrowseForFolder API function
Enum BrowseForFolderFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_EDITBOX = &H10
    BIF_RETURNFSANCESTORS = &H8
    BIF_USENEWUI = &H40
End Enum

'# consts used in browse for folder stuff
Private Const WM_USER            As Long = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const NOERROR            As Long = &H0&
Private Const S_OK               As Long = &H0&
Private Const S_FALSE            As Long = &H1&

'# type used with the SHBrowseForFolder API call
Private Type BrowseInfoType
     hwndOwner As Long
     pidlRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Function BrowseForFolderEx(hwnd As Long, Optional DlgTitle As String, Optional Flags As BrowseForFolderFlags, Optional SelectedPath As String) As String
Dim BFF_IT As BrowseInfoType
Dim itemID As Long
Dim selectedPathPointer As Long
Dim tmpPath As String * 256
Dim lR As Long
Dim LPTR As Long

    With BFF_IT
        .hwndOwner = hwnd                                   '# handle of the owner window
        If Len(DlgTitle) > 0 Then                           '# dialog text
            .lpszTitle = lstrcat(DlgTitle, "")              '# turn it into a number
        Else
            .lpszTitle = lstrcat("", "")
        End If
        If Len(SelectedPath) > 0 Then                       '# do we want to preselect a folder
            .lpfnCallback = FunctionPointer(AddressOf BrowseCallbackProcStr)    '# Dialog callback function that preselectes the folder specified
            selectedPathPointer = LocalAlloc(LPTR, Len(SelectedPath) + 1)       '# Allocate a string
            CopyMemory ByVal selectedPathPointer, ByVal SelectedPath, Len(SelectedPath) + 1     '# Copy the path to the string
            .lParam = selectedPathPointer                   '# The folder to preselect
            lR = CoInitialize(0)                            '# can we use bif_usenewui
            If (lR <> NOERROR) And (lR <> S_FALSE) Then
                .ulFlags = Flags                            '# just do default
            Else
                .ulFlags = Flags + BIF_USENEWUI             '# use it
            End If
        Else
            .ulFlags = Flags                                '# just do default
        End If
    End With
    
    itemID = SHBrowseForFolder(BFF_IT)                      '# Execute the BrowseForFolder API
    
    If itemID Then                                          '# path returned
        If SHGetPathFromIDList(itemID, tmpPath) Then        '# Get the path for the selected folder in the dialog
            BrowseForFolderEx = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) '# Take only the path without the nulls
        Else
            '# error occurred or invalid folder
            BrowseForFolderEx = "999Errors"
        End If
        Call CoTaskMemFree(itemID)                          '# Free the itemID
    Else
        '# cancel was pressed
        BrowseForFolderEx = "999Cancel"
    End If
    Call LocalFree(selectedPathPointer)                     '# Free the string from the memory

End Function

Private Function BrowseCallbackProcStr(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
'# processes messages for the dialog

    If uMsg = 1 Then
        Call SendMessage(hwnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If

End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long
'# returns the address of the callback function as a long

    FunctionPointer = FunctionAddress

End Function
