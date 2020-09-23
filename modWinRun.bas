Attribute VB_Name = "modWinRun"
Option Explicit

Public Sub RunExeAtStartup(Optional Arguments As String)
'# this function does the necessary to run the program at Windows startup with an option to randomize as well
Dim WshShell As Object
Dim strStartup As String
Dim objShellLink As Object
    
    On Local Error GoTo Err_Handler
    
    DoNotRunExeAtStartup                                                                '# delete any existing shortcut
    Set WshShell = CreateObject("WScript.Shell")                                        '# get reference to Wscript
    strStartup = WshShell.specialfolders("Startup")                                     '# get path to Windows startup folder
    Set objShellLink = WshShell.CreateShortcut(strStartup & "\" & App.EXEName & ".lnk") '# create the shortcut
    objShellLink.TargetPath = App.Path & "\" & App.EXEName & ".exe"                     '# set up target path
    If Arguments <> "" Then
        objShellLink.Arguments = Arguments                                              '# set up command line parameters
    End If
    objShellLink.WindowStyle = 1                                                        '# set to run in normal window
    objShellLink.IconLocation = App.Path & "\" & App.EXEName & ".exe, 0"                '# set up the icon to use
    objShellLink.WorkingDirectory = App.Path                                            '# set the working directory
    objShellLink.Save                                                                   '# save the shortcut
    Set WshShell = Nothing                                                              '# free up object
    Set objShellLink = Nothing                                                          '# free up object
    
    Exit Sub
    
Err_Handler:
    ErrPosition = "RunAtStartup()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub

Public Sub DoNotRunExeAtStartup()
'# this does the necessary to stop running
Dim fso As FileSystemObject
Dim WshShell As Object
Dim strStartup As String

    On Local Error GoTo Err_Handler
    
    Set fso = New FileSystemObject                                                      '# get fso object reference
    Set WshShell = CreateObject("WScript.Shell")                                        '# get reference to wscript
    strStartup = WshShell.specialfolders("Startup")                                     '# get path to windwos startup folder
    If fso.FileExists(strStartup & "\" & App.EXEName & ".lnk") Then                     '# check for existence of shorcut
        fso.DeleteFile (strStartup & "\" & App.EXEName & ".lnk")                        '# if exist delete it
    End If
    Set fso = Nothing                                                                   '# free up object
    Set WshShell = Nothing                                                              '#free up object
    DoEvents
    
    Exit Sub
    
Err_Handler:
    ErrPosition = "DoNotRunAtStartup()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    
End Sub
