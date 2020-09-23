Attribute VB_Name = "modPopupMenu"
Option Explicit

'# menu api functions/subs
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As PointApi)
Private Declare Sub SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long)
Private Declare Sub GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT)

'# consts for menu manipulation
Public Const MF_STRING          As Long = &H0
Public Const MF_GRAYED          As Long = &H1
Public Const MF_DISABLED        As Long = &H2
Public Const MF_CHECKED         As Long = &H8
Public Const MF_POPUP           As Long = &H10
Public Const MF_MENUBARBREAK    As Long = &H20
Public Const MF_OWNERDRAW       As Long = &H100
Public Const MFT_RADIOCHECK     As Long = &H200
Public Const MF_SEPARATOR       As Long = &H800
Public Const MF_BYCOMMAND       As Long = &H0&
Public Const MF_BYPOSITION      As Long = &H400&
Public Const MF_BITMAP          As Long = &H4&
Public Const MIIM_DATA          As Long = &H20

'# consts for menu tracking
Private Const TPM_RETURNCMD      As Long = &H100
Private Const TPM_LEFTALIGN      As Long = &H0&
Private Const TPM_RIGHTALIGN     As Long = &H8&

'# const used in graphics functions
Private Const SRCCOPY = &HCC0020

'# types used in the menu procedures
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type
Private Rct As RECT

Private Type MEASUREITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemWidth   As Long
    itemHeight  As Long
    itemData    As Long
End Type
Private MIS As MEASUREITEMSTRUCT

Private Type DRAWITEMSTRUCT
    CtlType     As Long
    CtlID       As Long
    itemID      As Long
    itemAction  As Long
    itemState   As Long
    hWndItem    As Long
    hdc         As Long
    rcItem      As RECT
    itemData    As Long
End Type
Private DIS As DRAWITEMSTRUCT

Private Type PointApi
    x           As Long
    y           As Long
End Type
Private Pt As PointApi

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private MII As MENUITEMINFO

'# used when drawing the menu
Private TotalMenuHeight As Long
Private intLoop As Integer
Public hMenu As Long
Public NumSeps As Long
Public FirstSep As Long

Public MenuIsOpen As Boolean

Public Sub DrawMenu(Frm As Form, ByVal hMenu As Long, ByRef PtrToDIS As Long, SidebarText As String, ByVal NumSeps As Long, ByVal FirstSep As Long)
'# does the actual drawing of the systray popup menu
Dim lngLoop As Long

    On Error GoTo Err_Handler
    
    CopyMemory DIS, ByVal PtrToDIS, Len(DIS)                        '# get drawing item structure
    With Frm
        Select Case DIS.itemID
            Case -1                                                 '# wants me to draw the sidebar we know because we set it in the call
                GetMenuItemRect .hwnd, hMenu, 1&, Rct               '# assuming item 1 is not a separator - get rect size/position of first menu item
                '# determine height of menu items
                TotalMenuHeight = (Rct.Bottom - Rct.Top) * (GetMenuItemCount(hMenu) - NumSeps - 1)
                If NumSeps Then '# if there are any seperators
                    GetMenuItemRect .hwnd, hMenu, FirstSep, Rct     '# get rect size/position of first seperator item
                    '# determine overall height of the menu
                    TotalMenuHeight = TotalMenuHeight + (Rct.Bottom - Rct.Top) * NumSeps
                End If
                With .picMenuCapt
                    .Move 0, 0, TotalMenuHeight, DIS.rcItem.Right   '# adjust size of picbox to fit sidebar
                    .BackColor = vbBlue                             '# set the default backcolor
                    .ForeColor = SBTEXTCOL                          '# set the colour of the text
                    .Cls                                            '# clear the box
                    For lngLoop = .ScaleWidth To 0 Step -1          '# lets have a gradient sidebar
                        frmMain.picMenuCapt.Line (lngLoop - 1, -1)-(lngLoop + 1, .ScaleWidth), RGB(REDVAL + ((lngLoop * 255) \ .ScaleWidth), GREENVAL + ((lngLoop * 255) \ .ScaleWidth), BLUEVAL + ((lngLoop * 255) \ .ScaleWidth)) '# mess with the rgb colours for different effects
                    Next lngLoop
                    With .Font                                      '# set the sidebar font
                        .Name = "Arial Black"
                        .Size = 10
                    End With '.FONT
                    .CurrentX = 10                                  '# placement of text in Sidebar - x pixels in
                    .CurrentY = 0
                    Frm.picMenuCapt.Print SidebarText               '# put the text on
                    '# rotated version of menu sidebar - done because setpixel quite slow so do the rotate before menu pops up
                    Frm.picMenuRot.Move 0, 0, DIS.rcItem.Right, TotalMenuHeight
                    With Pt
                        For .x = 0 To TotalMenuHeight - 1       '# this rotates pixelwise ccw 90°
                            For .y = 0 To DIS.rcItem.Right - 1
                                SetPixelV Frm.picMenuRot.hdc, .y, TotalMenuHeight - .x, Frm.picMenuCapt.Point(.x, .y)    '# rotates menu sidebar onto picMenuRot
                            Next .y
                        Next .x
                    End With 'PT
                    BitBlt DIS.hdc, 0, 0, DIS.rcItem.Right, TotalMenuHeight, Frm.picMenuRot.hdc, 0, 0, SRCCOPY
                End With '.PICMENUCAPT
            Case Else:
        End Select
    End With 'FRM

    Exit Sub
    
Err_Handler:
    ErrPosition = "DrawMenu()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Public Sub MeasureMenu(ByRef PtrToMIS As Long)
'# returns the dimensions of the menu items

    On Error GoTo Err_Handler
    
    CopyMemory MIS, ByVal PtrToMIS, Len(MIS)                '# copy memory into struct
    If MIS.itemID = -1& Then                                '# we are only interested in the sidebar
        MIS.itemWidth = 6                                   '# set width of sidebar, 6 seems OK
        CopyMemory ByVal PtrToMIS, MIS, Len(MIS)            '# copy struct back to memory
    End If

    Exit Sub
    
Err_Handler:
    ErrPosition = "MeasureMenu()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Public Sub TrackMenu(Frm As Form, hMenu As Long)
'# this function tracks the menu when it pops up

    On Error GoTo Err_Handler
    
    GetCursorPos Pt                                         '# get cursor position
    Frm.MenuClicked = TrackPopupMenu(hMenu, TPM_RETURNCMD Or TPM_LEFTALIGN, Pt.x, Pt.y, 0&, Frm.hwnd, 0&)

    Exit Sub
    
Err_Handler:
    ErrPosition = "TrackMenu()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub

Public Function GetMenuCaption(ByRef hMenu As Long, ByRef mItem As Long) As String
'# returns the caption of the menu item as a string for matching to the listview
Dim strCaption As String
Dim lngRetVal As Long
Dim intMax As Integer

    On Error GoTo Err_Handler
    
    intMax = 255
    strCaption = String(intMax, " ")                        '#set up a buffer
    lngRetVal = GetMenuString(hMenu, mItem, strCaption, intMax, MF_BYCOMMAND) '# get the caption off the menu
    strCaption = Left(strCaption, lngRetVal)                '# drop the null terminator
    GetMenuCaption = strCaption
    
    Exit Function

Err_Handler:
    ErrPosition = "GetMenuCaption()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Function
