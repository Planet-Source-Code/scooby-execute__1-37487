VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eXeCute!"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   178
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAbout 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picRestore 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2310
      Picture         =   "frmMain.frx":06EC
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picExit 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1860
      Picture         =   "frmMain.frx":0996
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picShortCut 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1410
      Picture         =   "frmMain.frx":0C40
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   -15
      Top             =   -15
   End
   Begin VB.PictureBox picMenuRot 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   945
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picMenuCapt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   480
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame Frame6 
      Height          =   1275
      Left            =   5603
      TabIndex        =   15
      Top             =   1305
      Width           =   1350
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   390
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "About eXeCute"
         Top             =   255
         Width           =   1050
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   390
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit eXeCute"
         Top             =   720
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options"
      Height          =   1275
      Left            =   195
      TabIndex        =   8
      Top             =   1305
      Width           =   4665
      Begin VB.Frame Frame5 
         Caption         =   "Use Tray Tips"
         Height          =   915
         Left            =   3105
         TabIndex        =   13
         Top             =   225
         Width           =   1365
         Begin VB.CheckBox chkUseTrayTips 
            Height          =   255
            Left            =   585
            TabIndex        =   14
            ToolTipText     =   "Check box to allow balloon style tooltips in the system tray"
            Top             =   405
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Run At Startup"
         Height          =   915
         Left            =   1635
         TabIndex        =   11
         Top             =   225
         Width           =   1365
         Begin VB.CheckBox chkRunAtStartup 
            Height          =   255
            Left            =   585
            TabIndex        =   12
            ToolTipText     =   "Check box to allow eXeCute to run at Windows startup"
            Top             =   390
            Width           =   195
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Play Sounds"
         Height          =   915
         Left            =   165
         TabIndex        =   9
         Top             =   225
         Width           =   1365
         Begin VB.CheckBox chkPlaySounds 
            Height          =   255
            Left            =   540
            TabIndex        =   10
            ToolTipText     =   "Check box to allow sounds to be played"
            Top             =   390
            Width           =   195
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Shortcut Folder Path"
      Height          =   1035
      Left            =   188
      TabIndex        =   5
      Top             =   210
      Width           =   6765
      Begin VB.CommandButton cmdBrowseForSCFolder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   165
         Picture         =   "frmMain.frx":0EEA
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Click to select a new shortcut folder"
         Top             =   300
         Width           =   600
      End
      Begin VB.ListBox lstShortcutFolder 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   240
         Left            =   915
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   5565
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPlaySounds_Click()

    PlayThisSound STSND_CLICK               '# play the corresponding sound
    PlaySounds = chkPlaySounds.Value        '# store the value in variable
    
End Sub

Private Sub chkRunAtStartup_Click()

    PlayThisSound STSND_CLICK               '# play the corresponding sound
    RunAtStartup = chkRunAtStartup.Value    '# store the value in variable
    If RunAtStartup Then                    '# if we are setting run at startup to true
        RunExeAtStartup                     '# put the shortcut in the startup folder
    Else                                    '# else
        DoNotRunExeAtStartup                '# delete the shortcut
    End If
    
End Sub

Private Sub chkUseTrayTips_Click()
    
    PlayThisSound STSND_CLICK               '# play the corresponding sound
    UseTips = chkUseTrayTips.Value
    
End Sub

Private Sub cmdBrowseForSCFolder_Click()
Dim strRetVal As String
    
    On Error GoTo Err_Handler
    
    PlayThisSound STSND_CLICK               '# play the corresponding sound
    
    '# do the Dialog bit
    strRetVal = BrowseForFolderEx(Me.hwnd, "Choose a folder:", BIF_DONTGOBELOWDOMAIN + BIF_RETURNONLYFSDIRS, ShortcutFolder)
    
    '# what was returned
    Select Case strRetVal
        Case "999Cancel":
                '# ignore cancel
        Case "999Errors":
                '# possible selected My Computer or some other invalid folder
            Err.Raise 40002
        Case ShortcutFolder:
                '# selected the same folder as currently selected
        Case Else:
            If strRetVal <> "" Then
                '# new folder selected
                ShortcutFolder = strRetVal
                lstShortcutFolder.Clear
                lstShortcutFolder.AddItem ShortcutFolder
            Else
                '# no folder selected
                Err.Raise 40003
            End If
    End Select
    
TidyUp:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 40002:                             '# invalid shorcut folder path
            MsgBox "You attempted to select an invalid Shortcut Folder" & vbCrLf & "Note : My Computer is NOT a folder", vbOKOnly + vbInformation, App.EXEName & " - Invalid Shortcut Folder"
        Case 40003:                             '# no folder selected
            MsgBox "You must select a Shortcut Folder", vbOKOnly + vbInformation, App.EXEName & " - Empty Shortcut Folder Path"
        Case Else:                              '# other errors
            ErrPosition = "cmdBrowseForSCFolder_Click()"
            ErrNumber = Err.Number
            ErrDesc = Err.Description
            ErrSource = Err.Source
            Call ErrorFunc
    End Select
    Resume TidyUp
    
End Sub

Private Sub Form_Load()

    On Error GoTo Err_Handler
    
    LoadingForm = True                          '# we don't play sounds whilst form is loading
    
    gHW = Me.hwnd                               '# get application handle
    Hook                                        '# take control of messages
    
    '# get path to Shortcut folder - default to \Shortcuts subdirectory under app.path
    ShortcutFolder = GetSetting(App.EXEName, "Settings", "ShortcutFolderPath", App.Path & "\Shortcuts")
    lstShortcutFolder.AddItem ShortcutFolder    '# put it in the list box
    
    '# check setting for playing sounds
    PlaySounds = GetSetting(App.EXEName, "Settings", "PlaySounds", 1)
    chkPlaySounds.Value = PlaySounds
    
    '# check setting for running prog at windows startup
    RunAtStartup = GetSetting(App.EXEName, "Settings", "RunAtStartup", 1)
    chkRunAtStartup.Value = RunAtStartup
    
    '# check setting for showing ballon tooltips
    UseTips = GetSetting(App.EXEName, "Settings", "UseTips", 1)
    chkUseTrayTips.Value = UseTips
    
    DisableX Me                                 '# disables th 'X' in the forms system menu
    Me.Show                                     '# show the form
    tmrHide.Enabled = True                      '# send the form to the systray in 5..4..3..2..1..
    
TidyUp:
    LoadingForm = False
    Exit Sub

Err_Handler:
    ErrPosition = "Form_Load()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    Resume TidyUp
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim Result As Long
Dim Msg As Long
       
    On Error GoTo Err_Handler
    
    '# the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        Msg = x
    Else
        Msg = x / Screen.TwipsPerPixelX
    End If
       
    Select Case Msg
        Case WM_LBUTTONUP:                      '# 514 restore form window
            RestoreForm                         '# display the main form
        Case WM_LBUTTONDBLCLK:                  '# 515 restore form window
            RestoreForm                         '# display the main form
        Case WM_RBUTTONDOWN:                    '# 516 display popup menu
            Result = SetForegroundWindow(Me.hwnd)   '# set the menu form as foreground
            Cr8Menu                             '# fill in the menu (redone every time so it picks up any changes to the directory structure)
            If Not MenuIsOpen Then
                MenuIsOpen = True               '# set the flag
                TrackMenu Me, hMenu             '# wait for menu messages
            End If
        Case Else:
    End Select
  
    Exit Sub
    
Err_Handler:
    ErrPosition = "Form_MouseMove()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Private Sub Form_Paint()
'# set the form colours to my personal choice

    ColourIn Me
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'    If UnloadMode = 0 Then     '# alternative method for disabling 'X' button but
'        Cancel = True          '# it doesn't remove from or grey out system menu items
'    End If

End Sub

Private Sub Form_Resize()
'# this is necessary to ensure that the minimized window is hidden
    
    On Error GoTo Err_Handler
    
    If Me.WindowState = vbMinimized Then        '# is window being minimized
        PlayThisSound STSND_MINIMIZE            '# play corresponding sound
        Me.Hide                                 '# hide the window
        '# put the icon in the system tray
        If UseTips Then
            AddToTray Me.hwnd, Me.Icon, "exeCute - Activate your shortcuts from the System Tray", App.TITLE, "Right_Click To Show Menu" & vbCrLf & "Left_Click To Show Form", ICO_INFO, 2000
        Else
            AddToTray Me.hwnd, Me.Icon, "exeCute - Activate your shortcuts from the System Tray"
        End If
    End If

    Exit Sub
    
Err_Handler:
    ErrPosition = "Form_Resize()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    On Error Resume Next
    
    SaveSetting App.EXEName, "Settings", "ShortcutFolderPath", ShortcutFolder   '# save the folder path
    SaveSetting App.EXEName, "Settings", "PlaySounds", PlaySounds               '# save the playsounds value
    SaveSetting App.EXEName, "Settings", "RunAtStartup", RunAtStartup           '# save the runatstartup value
    SaveSetting App.EXEName, "Settings", "UseTips", UseTips                     '# save the runatstartup value
    Unhook                                                                      '# return control of messages to Windows
    DestroyMenu hMenu                                                           '# destroy the shortcut menu you created
    RemoveFromTray                                                              '# remove the icon from the system tray

End Sub

Public Property Let MenuClicked(Item As Long)
Dim strShell As String
Dim strMenuCaption As String
Dim lngRet As Long

    MenuIsOpen = False
    
    '# update menu flags
    Select Case Item
        Case 0&:                                '# not clicked on valid item on the menu
        Case 997&:                              '# about option
            RestoreForm
            cmdAbout_Click
        Case 998&:                              '# restore option
            RestoreForm
        Case 999&:                              '# exit option
            Unload Me                           '# unload app
            Set frmMain = Nothing               '# clean up
        Case Else:                              '# clickable menu item
            strMenuCaption = GetMenuCaption(hMenu, Item)                                    '# get the menu caption (filename)
            strShell = Chr(34) & ShortcutFolder & "\" & strMenuCaption & ".lnk" & Chr(34)   '# create shell string
            LaunchLink Me, strShell                                                         '# launch the shortcut
    End Select

End Property

Private Sub RestoreForm()
'# called when the user clicks the popup menu Restore command
Dim Result As Long

    On Error GoTo Err_Handler
    
    PlayThisSound STSND_RESTORE                 '# play corresponding sound
    
    Me.WindowState = vbNormal                   '# set the window state to normal
    Result = SetForegroundWindow(Me.hwnd)       '# set the window as the one on the top
    Me.Show                                     '# show the form
    RemoveFromTray                              '# remove systray icon
    
    Exit Sub
    
Err_Handler:
    ErrPosition = "RestoreForm()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Private Sub cmdAbout_Click()
'# show the about form

    frmAbout.Show vbModal, Me
    
End Sub

Private Sub cmdExit_Click()
'# exit

    Unload Me
    Set frmMain = Nothing
    
End Sub

Private Sub lstShortcutFolder_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim lngXPoint As Long
Dim lngYPoint As Long
Dim lngIndex As Long

    On Error GoTo Err_Handler
    
    If Button = 0 Then                          '# if no button was pressed - mouseover
    lngXPoint = CLng(x / Screen.TwipsPerPixelX) '# get x position
    lngYPoint = CLng(y / Screen.TwipsPerPixelY) '# get y position
        With lstShortcutFolder
            '# get selected item from list (there is only one)
            lngIndex = SendMessage(.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((lngYPoint * 65536) + lngXPoint))
            '# show tip or clear last one
            If (lngIndex >= 0) And (lngIndex <= .ListCount) Then
                .ToolTipText = .List(lngIndex)
            Else
                .ToolTipText = ""
            End If
        End With
    End If
    
    Exit Sub
        
Err_Handler:
    ErrPosition = "tmrHide_Timer()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub

Private Sub tmrHide_Timer()
'# hides the form after a specified period of time
    
    On Error GoTo Err_Handler
    
    If Not FrmAboutOpen Then
        tmrHide.Enabled = False                     '# disable the timer
        Me.WindowState = vbMinimized                '# minimize the form
    End If
    
    Exit Sub
        
Err_Handler:
    ErrPosition = "tmrHide_Timer()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub

Private Sub Cr8Menu()
'# does the creating of the popup menu for the systray
Dim strCaption As String
Dim arrCaption() As String
Dim intMax As Integer
Dim lngCurrMenu As Long
Dim lngMenuPos As Long
Dim oFso As FileSystemObject
Dim oFolder As Folder
Dim oFile As File
Dim intLoop As Integer

    On Error GoTo Err_Handler
    
    '# set the number of seperators (1 between shortcuts and restore/exit options)
    NumSeps = 2

    '# delete the current menu
    DestroyMenu hMenu
    
    '# create a new menu
    hMenu = CreatePopupMenu
    
    '# draw the sidebar - Item Id (-1) to recognize the sidebar
    AppendMenu hMenu, MF_OWNERDRAW Or MF_DISABLED, -1&, 0&
    
    '# create each main menu item i.e. each shortcut
    intLoop = 0
    Set oFso = New FileSystemObject
    If oFso.FolderExists(ShortcutFolder) Then                   '# check the shorcut folder exists
        Set oFolder = oFso.GetFolder(ShortcutFolder)            '# get a reference to it
        If oFolder.Files.Count > 0 Then                         '# check there are file items in the folder
            For Each oFile In oFolder.Files                     '# for each file item
                If LCase(Right(oFile.Name, 4)) = ".lnk" Then    '# check it's a link
                    intLoop = intLoop + 1                       '# increment the counter
                    strCaption = Left(oFile.Name, Len(oFile.Name) - 4) & vbNullChar '# set the menu caption
                    If intLoop = 1 Then                         '# if it's the first item we need another column
                        AppendMenu hMenu, MF_MENUBARBREAK, intLoop, strCaption  '# add to menu
                    Else
                        AppendMenu hMenu, 0&, intLoop, strCaption               '# add to the menu
                    End If
                    '# add our shortcut picture
                    SetMenuItemBitmaps hMenu, intLoop, MF_BITMAP, picShortCut.Picture, picShortCut.Picture
                End If
            Next
        End If
    Else
        Err.Raise 40001
    End If
    Set oFolder = Nothing                       '# free up objects
    Set oFso = Nothing
    
    If intLoop = 0 Then                         '# if no shortcut items
        intLoop = 1
        AppendMenu hMenu, MF_GRAYED Or MF_DISABLED Or MF_MENUBARBREAK, 0&, "No shortcuts found ..."
    End If
    
    '# create a separator
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    
    '# position of the seperator on the menu
    FirstSep = intLoop + 1                      '# i.e. if 4 shortcuts then FirstSep = 5
    
    '# create static menu items and pictures
    AppendMenu hMenu, 0&, 997&, "About"
    SetMenuItemBitmaps hMenu, 997&, MF_BITMAP, picAbout.Picture, picAbout.Picture
    
    AppendMenu hMenu, 0&, 998&, "Restore"
    SetMenuItemBitmaps hMenu, 998&, MF_BITMAP, picRestore.Picture, picRestore.Picture
    
    '# create a separator
    AppendMenu hMenu, MF_SEPARATOR, 0&, 0&
    
    AppendMenu hMenu, 0&, 999&, "Exit"
    SetMenuItemBitmaps hMenu, 999&, MF_BITMAP, picExit.Picture, picExit.Picture
        
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
        Case 40001:                             '# invalid shorcut folder path
            MsgBox "You must select a valid Shortcut Folder", vbOKOnly + vbInformation, App.EXEName & " - Invalid Shortcut Folder"
            RestoreForm
            cmdBrowseForSCFolder_Click
        Case Else
            ErrPosition = "Cr8Menu()"
            ErrNumber = Err.Number
            ErrDesc = Err.Description
            ErrSource = Err.Source
            Call ErrorFunc
    End Select

End Sub
