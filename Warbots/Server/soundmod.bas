Attribute VB_Name = "soundmod"
Option Explicit

'*-------------------------------------*
'* Playsound flags: store in dwFlags   *
'*-------------------------------------*
' lpszName points to a registry entry
' Do not use SND_RESOURSE or SND_FILENAME
Private Const SND_ALIAS& = &H10000
' Playsound returns immediately
' Do not use SND_SYNC
Private Const SND_ASYNC& = &H1
' The name of a wave file.
' Do not use with SND_RESOURCE or SND_ALIAS
Private Const SND_FILENAME& = &H20000
' Unless used, the default beep will
' play if the specified resource is missing
Private Const SND_NODEFAULT& = &H2
' Fail the call & do not wait for
' a sound device if it is otherwise unavailable
Private Const SND_NOWAIT& = &H2000
' Use a resource file as the source.
' Do not use with SND_ALIAS or SND_FILENAME
Private Const SND_RESOURCE& = &H40004
' Playsound will not return until the
' specified sound has played.  Do not
' use with SND_ASYNC
Private Const SND_SYNC& = &H0

Public Enum enSound_Source
    ssFile = SND_FILENAME&
    ssRegistry = SND_ALIAS&
End Enum

' These are common sounds available from the registry
Public Const elDefault = ".Default"
Public Const elGPF = "AppGPFault"
Public Const elClose = "Close"
Public Const elEmptyRecycleBin = "EmptyRecycleBin"
Public Const elMailBeep = "MailBeep"
Public Const elMaximize = "Maximize"
Public Const elMenuCommand = "MenuCommand"
Public Const elMenuPopUp = "MenuPopup"
Public Const elMinimize = "Minimize"
Public Const elOpen = "Open"
Public Const elRestoreDown = "RestoreDown"
Public Const elRestoreUp = "RestoreUp"
Public Const elSystemAsterisk = "SystemAsterisk"
Public Const elSystemExclaimation = "SystemExclaimation"
Public Const elSystemExit = "SystemExit"
Public Const elSystemHand = "SystemHand"
Public Const elSystemQuestion = "SystemQuestion"
Public Const elSystemStart = "SystemStart"

Private Declare Function PlaySound& Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long)
' hModule is only used if SND_RESOURCE& is set and represents
' an HINSTANCE handle.  This example doesn't support playing
' from a resource file.

' Plays sounds from the registry or a disk file
' Doesn't care if the file is missing
Public Function EZPlay(ssname As String, _
    sound_source As enSound_Source) As Boolean
   
    If PlaySound(ssname, 0&, sound_source + _
        SND_ASYNC + SND_NODEFAULT) Then
        EZPlay = True
    Else
        EZPlay = False
    End If
   
End Function

