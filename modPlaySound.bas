Attribute VB_Name = "modPlaySound"
Option Explicit

'# api function to play a sound
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Byte, ByVal uFlags As Long) As Long

'# consts for playing sounds
Private Const SND_ASYNC          As Long = &H1       '# play asynchronously
Private Const SND_NODEFAULT      As Long = &H2       '# do not play the default sound
Private Const SND_MEMORY         As Long = &H4       '# sound is buffered in memory
Private Const SND_FILENAME       As Long = &H20000   '# name is a file name
Private Const SND_LOOP           As Long = &H8       '# loop the sound until next sndPlaySound
Private Const SND_NOWAIT         As Long = &H2000    '# don't wait if the driver is busy
Private Const SND_RESOURCE       As Long = &H40004   '# name is a resource name or atom
Private Const SND_SYNC           As Long = &H0       '# play synchronously (default)

'# byte array used to play sounds stored in resource files
Private bytSound() As Byte

'# enum for the sounds saved in resource file
Public Enum StoredSounds
    STSND_RESTORE = 101
    STSND_MINIMIZE = 101
    STSND_ERROR = 102
    STSND_LAUNCH = 103
    STSND_CLICK = 104
End Enum

Public Function LoadSoundFromRes(SoundID As Variant, Optional strType As String = "CUSTOM") As Boolean
'# loads a sound stored in a resource file
Dim bytArray() As Byte
Dim lngIndexCounter As Long
Dim lngFileLen As Long
  
    On Error GoTo Err_Handler
  
    LoadSoundFromRes = True
  
    bytArray = LoadResData(SoundID, strType)                                            '# load sound data into byte array
  
    lngFileLen = UBound(bytArray)                                                       '# get length of data
    ReDim bytSound(lngFileLen)                                                          '# redimension a global byte array to store the sound in
  
    For lngIndexCounter = 0 To lngFileLen
        bytSound(lngIndexCounter) = bytArray(lngIndexCounter)                           '# put data into array
    Next
  
    Exit Function
  
Err_Handler:
    ErrPosition = "LoadSoundFromRes()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc
    LoadSoundFromRes = False
  
End Function

Public Sub PlayThisSound(SoundID As StoredSounds)
Dim SoundLoadedOK As Boolean
    
    On Error GoTo Err_Handler
    
    If PlaySounds And Not LoadingForm Then                                      '# play sounds if we are allowed
        SoundLoadedOK = LoadSoundFromRes(SoundID, "SOUND")                      '# load the sound into byte array
        If SoundLoadedOK Then                                                   '# if it worked
            sndPlaySound bytSound(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY  '# play it
        End If
    End If

    Exit Sub
    
Err_Handler:
    ErrPosition = "PlayThisSound()"
    ErrNumber = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    Call ErrorFunc

End Sub
