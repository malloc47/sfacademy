Attribute VB_Name = "Modmain"
Option Explicit

'INI Stuff
Declare Function writeprivateprofilestring Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long

Declare Function getprivateprofilestring Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Dim sSpaces As String

Dim sSpaces2 As String

Dim sSpaces3 As String

Dim sSpaces4 As String

Dim sSpaces5 As String

Dim sSpaces6 As String

Dim szReturn As String

Dim i1 As Long

Dim i2 As Long

Dim i3 As Long

Dim i4 As Long

Dim i5 As Long

Dim i6 As Long

Dim CodesLong As Long

Dim MusicValue As Long

Dim SoundValue As Long

Dim MusicToggle As String

Dim SoundToggle As String

Dim sValue1 As String

Dim sValue2 As String

Dim sValue3 As String

Dim sValue4 As String

Dim sValue5 As String

Dim sValue6 As String

Dim CodeString As String

Dim MusValue As String

Dim used As Boolean

'End ini stuff

Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long

'Midi Sound Code
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBffer As String, ByVal uLength As Long) As Long
Declare Function GetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'Midi Sound Code End
Public WavFile As String

Public WaveCheck As Long

Public sfile As String

Public NUMCodes As String

Public Screen1 As Integer
Public Door1Open As Integer
Public Elevator1Open As Integer
Public Elevator1Closed As Boolean
Public Elevator1Count As Integer
Public ButtonClicked As Boolean
Public PipeTaken As Boolean
Public PipeAV As Boolean
Public PipeUsed As Boolean
Public Button2Clicked As Boolean
Public MusicNum As Integer
Public SongNum As Integer
Public MusicOn As Boolean
Public SoundOn As Boolean
Public ReSound As Boolean
Public DialogueToggle As Boolean
Public MessageBox As Long
Public ElevatorDoorsOpen As Boolean

Dim CDPlayer
'Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long
Dim rc
' constants for snyPlaySound
Const SND_SYNC = (0)
Const SND_ASYNC = (1)
Const SND_NODEFAULT = (2)
Const SND_MEMORY = (4)
Const SND_LOOP = (8)
Const SND_NOSTOP = (16)
Public Sub Slidesoundstart()
If SoundOn = True And ReSound = False Then rc = sndplaysound(App.Path & "\sound\slide.wav", SND_NODEFAULT + SND_ASYNC + SND_LOOP)
End Sub
Public Sub Pipeinv()
If PipeAV = True Then frminv.imgpipe.Visible = True
End Sub

Public Sub Playsoundclick()
'Dim rc
If SoundOn = True Then rc = sndplaysound(App.Path & "\sound\buttonclick.wav", SND_NODEFAULT + SND_ASYNC)
End Sub

Public Sub Playsoundpipetake()
If SoundOn = True Then rc = sndplaysound(App.Path & "\sound\pipetake.wav", SND_NODEFAULT + SND_ASYNC)
End Sub

Public Sub Playsoundunjam()
If SoundOn = True Then rc = sndplaysound(App.Path & "\sound\unjam.wav", SND_NODEFAULT + SND_ASYNC)
End Sub
Public Sub Moveleft()
If SoundOn = True And ReSound = False Then rc = sndplaysound(App.Path & "\sound\moveleft.wav", SND_NODEFAULT + SND_ASYNC)
End Sub

Public Sub PlaySound()
If SoundOn = True Then rc = sndplaysound(App.Path & WavFile, SND_NODEFAULT + SND_ASYNC)
End Sub
Public Sub PlaySoundLoop()
If SoundOn = True Then rc = sndplaysound(App.Path & WavFile, SND_NODEFAULT + SND_ASYNC + SND_LOOP)
End Sub


'More Midi Stuff

Public Sub OpenMidi()
 Dim sShortFile As String * 67
 Dim lResult As Long
 Dim sError As String * 255

 'The mciSendString API call doesn't seem to like
 'long filenames that have spaces in them, so we
 'will make another API call to get the short
 'filename version.

 lResult = GetShortPathName(sfile, sShortFile, Len(sShortFile))
 sfile = Left(sShortFile, lResult)

 'Make the call to open the midi file and assign 'it an alias
 lResult = mciSendString("open " & sfile & " type sequencer alias mcitest", ByVal 0&, 0, 0)

 'Check to see if there was an error
 If lResult Then
   lResult = mciGetErrorString(lResult, sError, 255)
   Debug.Print "open: " & sError
 End If
End Sub

Public Sub PlayMidi()
 Dim lResult As Integer
 Dim sError As String * 255

'Make the call to start playing the midi
 lResult = mciSendString("play mcitest", ByVal 0&, 0, 0)

'Check to see if there were any errors
 If lResult Then
   lResult = mciGetErrorString(lResult, sError, 255)
   Debug.Print "play: " & sError
 End If
End Sub

Public Sub CloseMidi()
 Dim lResult As Integer
 Dim sError As String * 255

'Make the call to close the midi file
 lResult = mciSendString("close mcitest", "", 0&, 0&)

'Check to see if there were any errors
 If lResult Then
   lResult = mciGetErrorString(lResult, sError, 255)
   Debug.Print "stop: " & sError
 End If
End Sub

Public Sub Speechoptionsdisabled()
frmdialogue.Option1.Enabled = False
frmdialogue.Option2.Enabled = False
frmdialogue.Option3.Enabled = False
frmdialogue.Option4.Enabled = False
frmdialogue.Cmdtalk.Enabled = False
End Sub

Public Sub OpenCDPlayer()
On Error GoTo CdPlayerError
CDPlayer = Shell("C:\windows\cdplayer.exe", 1)
CdPlayerError:
If Err.Number = 53 Then
MsgBox ("Could Not Find Windows CD Player!")
Exit Sub
Else

Exit Sub

End If

End Sub
Public Sub SongMnuClicked(Songargument As Integer)
CloseMidi
Clearsongmenu
MusicOn = True
frmimage.mnumusic.Checked = True
SongNum = Songargument
Songchange
frmimage.Timer1.Enabled = True
MusicNum = 1
End Sub

Public Sub Clearsongmenu()
frmimage.mnusti.Checked = False
frmimage.mnustii.Checked = False
frmimage.mnustiii.Checked = False
frmimage.mnustiv.Checked = False
frmimage.mnusttng.Checked = False
frmimage.mnustvi.Checked = False
frmimage.mnustvii.Checked = False
frmimage.mnustviii.Checked = False
frmimage.mnustds9.Checked = False
frmimage.mnustvoy.Checked = False
frmimage.mnusttos.Checked = False

End Sub

Public Sub Songchange()
If MusicOn = True Then
CloseMidi
Clearsongmenu

If SongNum = 1 Then
sfile = App.Path & "\sound\st_i.mid"
frmimage.mnusti.Checked = True

ElseIf SongNum = 2 Then
sfile = App.Path & "\sound\st_ii.mid"
frmimage.mnustii.Checked = True

ElseIf SongNum = 3 Then
sfile = App.Path & "\sound\st_iii.mid"
frmimage.mnustiii.Checked = True

ElseIf SongNum = 4 Then
sfile = App.Path & "\sound\st_iv.mid"
frmimage.mnustiv.Checked = True

ElseIf SongNum = 5 Then
sfile = App.Path & "\sound\st_tng.mid"
frmimage.mnusttng.Checked = True

ElseIf SongNum = 6 Then
sfile = App.Path & "\sound\st_vi.mid"
frmimage.mnustvi.Checked = True

ElseIf SongNum = 7 Then
sfile = App.Path & "\sound\st_vii.mid"
frmimage.mnustvii.Checked = True

ElseIf SongNum = 8 Then
sfile = App.Path & "\sound\st_viii.mid"
frmimage.mnustviii.Checked = True

ElseIf SongNum = 9 Then
sfile = App.Path & "\sound\st_ds9.mid"
frmimage.mnustds9.Checked = True

ElseIf SongNum = 10 Then
sfile = App.Path & "\sound\st_voy.mid"
frmimage.mnustvoy.Checked = True

ElseIf SongNum = 11 Then
sfile = App.Path & "\sound\st_tos.mid"
frmimage.mnusttos.Checked = True

End If
OpenMidi
PlayMidi
End If
End Sub

Public Sub MusicVarChange()
If MusicOn = True Then

MusicNum = MusicNum + 1
'frmdialogue.Text1.Text = Musicnum

If SongNum = 1 Then
    If MusicNum = 9 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 2 Then
    If MusicNum = 13 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 3 Then
    If MusicNum = 10 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 4 Then
    If MusicNum = 9 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 5 Then
    If MusicNum = 6 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 6 Then
    If MusicNum = 17 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 7 Then
    If MusicNum = 7 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 8 Then
    If MusicNum = 10 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 9 Then
    If MusicNum = 9 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 10 Then
    If MusicNum = 7 Then
        SongNum = SongNum + 1
        Songchange
        MusicNum = 1
    End If

ElseIf SongNum = 11 Then
    If MusicNum = 5 Then
        SongNum = 1
        Songchange
        MusicNum = 1
    End If
                
End If

'If Musicnum > 4 Then
'Musicnum = 1
'End If

End If

End Sub

Public Sub INIGet()
used = False

sSpaces = Space(75)

sSpaces2 = Space(75)

sSpaces3 = Space(75)

sSpaces4 = Space(75)

sSpaces5 = Space(75)

sSpaces6 = Space(250)
 
MusicToggle = Space(75)

SoundToggle = Space(75)

NUMCodes = Space(75)

getprivateprofilestring "options.main", "FTop", szReturn, sSpaces, Len(sSpaces), "sfacademy.ini"

getprivateprofilestring "options.main", "FLeft", szReturn, sSpaces2, Len(sSpaces2), "sfacademy.ini"

getprivateprofilestring "options.main", "Dtop", szReturn, sSpaces3, Len(sSpaces3), "sfacademy.ini"

getprivateprofilestring "options.main", "DLeft", szReturn, sSpaces4, Len(sSpaces4), "sfacademy.ini"

getprivateprofilestring "options.main", "ITop", szReturn, sSpaces5, Len(sSpaces5), "sfacademy.ini"

getprivateprofilestring "options.main", "ILeft", szReturn, sSpaces6, Len(sSpaces6), "sfacademy.ini"

getprivateprofilestring "sound.values", "music", szReturn, MusicToggle, Len(sSpaces6), "sfacademy.ini"

getprivateprofilestring "sound.values", "sound", szReturn, SoundToggle, Len(sSpaces6), "sfacademy.ini"

getprivateprofilestring "level.codes", "levelcodes", szReturn, NUMCodes, Len(sSpaces6), "sfacademy.ini"

On Error GoTo fixerror
frmimage.Top = sSpaces
frmimage.Left = sSpaces2
frmdialogue.Top = sSpaces3
frmdialogue.Left = sSpaces4
frminv.Top = sSpaces5
frminv.Left = sSpaces6

If MusicToggle = 1 Then
MusicChange
Else
    frmimage.mnumusic.Checked = False
    MusicOn = False
    frmimage.Timer1.Enabled = True
    CloseMidi
    frmimage.mnuchoosesong.Enabled = False
End If

If SoundToggle = 1 Then
SoundChange
ElseIf SoundToggle = 0 Then
    frmimage.mnusoundon.Checked = False
    SoundOn = False
End If

used = True

fixerror:
If used = False Then
frmimage.Left = 2610
frmimage.Top = 1065
frmdialogue.Left = 1065
frmdialogue.Top = 1065
frminv.Left = 2610
frminv.Top = 5325
NUMCodes = 1
Else

End If

End Sub

Public Sub INIMake()

sValue1 = frmimage.Top

sValue2 = frmimage.Left

sValue3 = frmdialogue.Top

sValue4 = frmdialogue.Left

sValue5 = frminv.Top

sValue6 = frminv.Left

i1 = writeprivateprofilestring("options.main", "FTop", sValue1, "sfacademy.ini")

i2 = writeprivateprofilestring("options.main", "FLeft", sValue2, "sfacademy.ini")

i3 = writeprivateprofilestring("options.main", "DTop", sValue3, "sfacademy.ini")

i4 = writeprivateprofilestring("options.main", "DLeft", sValue4, "sfacademy.ini")

i5 = writeprivateprofilestring("options.main", "ITop", sValue5, "sfacademy.ini")

i6 = writeprivateprofilestring("options.main", "ILeft", sValue6, "sfacademy.ini")

CodesLong = writeprivateprofilestring("level.codes", "levelcodes", NUMCodes, "sfacademy.ini")

If frmimage.mnumusic.Checked = True Then
MusicValue = writeprivateprofilestring("sound.values", "music", "1", "sfacademy.ini")
Else
MusicValue = writeprivateprofilestring("sound.values", "music", "0", "sfacademy.ini")
End If

If frmimage.mnusoundon.Checked = True Then
SoundValue = writeprivateprofilestring("sound.values", "sound", "1", "sfacademy.ini")
Else
SoundValue = writeprivateprofilestring("sound.values", "sound", "0", "sfacademy.ini")
End If

End Sub

Public Sub SoundChange()
    WaveCheck = waveOutGetNumDevs()
    If WaveCheck > 0 Then
        SoundOn = True
        frmimage.mnusoundon.Checked = True
    Else
        MsgBox ("Your system can't play sound files, sound has been disabled")
        SoundOn = False
        frmimage.mnusoundon.Checked = False
    End If

End Sub
Public Sub SoundQuit()
        SoundOn = False
        frmimage.mnusoundon.Checked = False
End Sub
Public Sub MusicChange()
WaveCheck = waveOutGetNumDevs()
    If WaveCheck > 0 Then
frmimage.mnumusic.Checked = True
MusicOn = True
frmimage.Timer1.Enabled = False
OpenMidi
PlayMidi
MusicNum = 1
SongNum = 1
frmimage.mnuchoosesong.Enabled = True
Songchange
Else
    MsgBox ("Your system can't play sound files, sound has been disabled.")
    frmimage.mnumusic.Checked = False
    MusicOn = False
    frmimage.Timer1.Enabled = True
    CloseMidi
    frmimage.mnuchoosesong.Enabled = False
End If
End Sub
Public Sub MusicQuit()
    frmimage.mnumusic.Checked = False
    MusicOn = False
    frmimage.Timer1.Enabled = True
    CloseMidi
    frmimage.mnuchoosesong.Enabled = False
End Sub
