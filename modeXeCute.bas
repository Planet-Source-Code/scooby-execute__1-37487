Attribute VB_Name = "modeXeCute"
Option Explicit

'# api function and const for the tooltip for the ShortcutFolder
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_ITEMFROMPOINT = &H1A9

'# my own consts used in the program
Public Const DEBUGGING_ON = True
Public Const WRITE_DBG = False
Public Const SBTEXT = "-- eXeCute --"
'# play with this to change the sidebar forecolour
Public Const SBTEXTCOL = vbBlack
'# play with these to change the sidebar gradient backcolour
Public Const REDVAL = 100
Public Const GREENVAL = 100
Public Const BLUEVAL = 100

'# variables used throughout the program
'# for error processing
Public ErrPosition As String
Public ErrNumber As Long
Public ErrDesc As String
Public ErrSource As String

'# path to the folder containing all the shorcuts
Global ShortcutFolder As String

'# to see if we are playing/able to play sounds
Global PlaySounds As Integer
'# to see if program should be run at windows startup
Global RunAtStartup As Integer
'# to see if we are using balloon tooltips
Global UseTips As Integer

'# to see if we are loading the main form
Global LoadingForm As Boolean
'# to say about form open
Global FrmAboutOpen As Boolean

Public Sub ErrorFunc()
'# builds the error message to output
    
    On Error Resume Next
        
    PlayThisSound STSND_ERROR       '# play the corresponding sound

    Dbg "Error " & ErrNumber & " [" & ErrDesc & "] has occurred in " & ErrPosition
    ErrNumber = 0                   '# reset error variables
    ErrDesc = ""
    ErrPosition = ""
    ErrSource = ""
    
End Sub

Private Sub Dbg(ByVal DbgMsg As String)
'# ouputs debug/error information either to file or to the screen (msgboxes)
'# DEBUGGING_ON and WRITE_DBG are booleans (recommend always True for DEBUGGING_ON
'# and False for WRITE_DBG while testing until you are happy that errors are catered
'# for then set to True so the user isn't annoyed by pointless MsgBoxes).
Dim intFF As Integer
Dim DbgFile As String
    
    If DEBUGGING_ON Then
        DbgMsg = Format$(Now, "dd/mm/yy hh:nn") & " - " & DbgMsg                    '# datestamp error message
        If WRITE_DBG Then
            On Local Error Resume Next                                              '# ignore errors
            intFF = FreeFile                                                        '# get next freefile number
            DbgFile = App.Path & App.EXEName & ".dbg"                               '# set the debug file path
            Open DbgFile For Append As #intFF                                       '# open for append
            Print #intFF, DbgMsg                                                    '# write the message
            Close #intFF                                                            '# close the file
        Else
            MsgBox DbgMsg, vbOKOnly + vbInformation, App.EXEName & " - Warning!"    '# show error
        End If
    End If
    DoEvents                                                                        '# be nice to windows
    
End Sub

Public Sub ColourIn(Frm As Form)
Dim ctrl As Control

    On Error Resume Next                        '# ignore controls that don't support the property
    Frm.BackColor = RGB(200, 200, 200)          '# do the form colours
    Frm.ForeColor = vbBlack
    For Each ctrl In Frm.Controls               '# do each control to match the form
        ctrl.BackColor = Frm.BackColor
        ctrl.ForeColor = Frm.ForeColor
    Next ctrl

End Sub
