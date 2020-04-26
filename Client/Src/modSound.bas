Attribute VB_Name = "modSound"
Option Explicit

Public Performance As DirectMusicPerformance
Public Segment As DirectMusicSegment
Public Loader As DirectMusicLoader

Public DS As DirectSound

Public Const SOUND_BUFFERS = 50

Private Type BufferCaps
    Volume As Boolean
    Frequency As Boolean
    Pan As Boolean
End Type

Private Type SoundArray
    DSBuffer As DirectSoundBuffer
    DSCaps As BufferCaps
    DSSourceName As String
End Type

Private Sound(1 To SOUND_BUFFERS) As SoundArray

' Contains the current sound index.
Public SoundIndex As Long

Public Music_On As Boolean
Public Music_Playing As String

Public Sound_On As Boolean
Private SEngineRestart As Boolean

Private Const DefaultVolume As Long = 100

Public Sub InitMusic()

    Set Loader = DX7.DirectMusicLoaderCreate
    Set Performance = DX7.DirectMusicPerformanceCreate
   
    Performance.Init Nothing, frmMainGame.hwnd
    Performance.SetPort -1, 80
   
    ' adjust volume 0-100
    Performance.SetMasterVolume DefaultVolume * 42 - 3000
    Performance.SetMasterAutoDownload True
   
End Sub

Public Sub InitSound()

    'Make the DirectSound object
    Set DS = DX7.DirectSoundCreate(vbNullString)
   
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    DS.SetCooperativeLevel frmMainGame.hwnd, DSSCL_PRIORITY
   
End Sub

Private Function GetState(ByVal Index As Integer) As String
    'Returns the current state of the given sound
    GetState = Sound(Index).DSBuffer.GetStatus
End Function

Public Sub SoundStop(ByVal Index As Integer)

    'Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
   
End Sub

Private Sub SoundLoad(ByVal file As String)
Dim DSBufferDescription As DSBUFFERDESC
Dim DSFormat As WAVEFORMATEX

    ' Set the sound index one higher for each sound.
    SoundIndex = SoundIndex + 1
   
    ' Reset the sound array if the array height is reached.
    If SoundIndex > UBound(Sound) Then
        SEngineRestart = True
        SoundIndex = 1
    End If
   
    ' Remove the sound if it exists (needed for re-loop).
    If SEngineRestart Then
        If GetState(SoundIndex) = DSBSTATUS_PLAYING Then
            SoundStop SoundIndex
            SoundRemove SoundIndex
        End If
    End If
   
    ' Load the sound array with the data given.
    With Sound(SoundIndex)
        .DSSourceName = file            'What is the name of the source?
        .DSCaps.Pan = True              'Is this sound to have Left and Right panning capabilities?
        .DSCaps.Volume = True           'Is this sound capable of altered volume settings?
    End With
   
    'Set the buffer description according to the data provided
    With DSBufferDescription
        If Sound(SoundIndex).DSCaps.Pan Then
            .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        End If
        If Sound(SoundIndex).DSCaps.Volume Then
            .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
        End If
    End With
   
    'Set the Wave Format
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 22050
        .nBitsPerSample = 16
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
   
    Set Sound(SoundIndex).DSBuffer = DS.CreateSoundBufferFromFile(App.Path & SOUND_PATH & Sound(SoundIndex).DSSourceName, DSBufferDescription, DSFormat)
   
End Sub

Public Sub SoundRemove(ByVal Index As Integer)
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing
        .DSCaps.Frequency = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSSourceName = vbNullString
    End With
End Sub

Private Sub SetVolume(ByVal Index As Integer, ByVal Vol As Long)
    'Check to make sure that the buffer has the capability of altering its volume
    If Not Sound(Index).DSCaps.Volume Then Exit Sub

    'Alter the volume according to the Vol provided
    If Vol > 0 Then
        Sound(Index).DSBuffer.SetVolume (60 * Vol) - 6000
    Else
        Sound(Index).DSBuffer.SetVolume -6000
    End If
End Sub

Private Sub SetPan(ByVal Index As Integer, ByVal Pan As Long)
    'Check to make sure that the buffer has the capability of altering its pan
    If Not Sound(Index).DSCaps.Pan Then Exit Sub

    'Alter the pan according to the pan provided
    Select Case Pan
        Case 0
            Sound(Index).DSBuffer.SetPan -10000
        Case 100
            Sound(Index).DSBuffer.SetPan 10000
        Case Else
            Sound(Index).DSBuffer.SetPan (100 * Pan) - 5000
    End Select
End Sub

Public Sub PlayMidi(ByVal FileName As String, ByVal repeats As Long)
Dim Splitmusic() As String

    Splitmusic = Split(FileName, ".", , vbTextCompare)
   
    If Performance Is Nothing Then Exit Sub
    If LenB(Trim$(FileName)) < 1 Then Exit Sub
    If UBound(Splitmusic) <> 1 Then Exit Sub
    If Splitmusic(1) <> "mid" Then Exit Sub
    If Not FileExist(App.Path & MUSIC_PATH & FileName, True) Then Exit Sub
   
    If Not Music_On Then Exit Sub
   
    If Music_Playing = FileName Then Exit Sub
   
    Set Segment = Nothing
    Set Segment = Loader.LoadSegment(App.Path & MUSIC_PATH & FileName)
   
    ' repeat midi file
    Segment.SetLoopPoints 0, 0
    Segment.SetRepeats 100
    Segment.SetStandardMidiFile
   
    Performance.PlaySegment Segment, 0, 0
   
    Music_Playing = FileName
   
End Sub

Public Sub StopMidi()
    If Not (Performance Is Nothing) Then Performance.Stop Segment, Nothing, 0, 0
    Music_Playing = vbNullString
End Sub

Public Sub PlaySound(ByVal file As String, Optional ByVal Volume As Long = 100, Optional ByVal Pan As Long = 50)
     
    ' Check to see if DirectSound was successfully initalized.
    If Not Sound_On Or Not FileExist(App.Path & SOUND_PATH & file, True) Then Exit Sub
    
    If Options.PlayMusic = 0 Then
    If file = "Choose.wav" Then
    Else
    Exit Sub
    End If
    End If
    
    ' Loads our sound into memory.
    SoundLoad file
   
    ' Sets the volume for the sound.
    SetVolume SoundIndex, Volume
   
    ' Sets the pan for the sound.
    SetPan SoundIndex, Pan
   
    ' Play the sound.
    
    Sound(SoundIndex).DSBuffer.Play DSBPLAY_DEFAULT
    
End Sub

Function GetSoundPosition(ByVal Index As Long) As Long
Dim curr As DSCURSORS

GetSoundPosition = curr.lWrite
End Function

Public Sub PlayMusic(ByVal file As String)
PlaySound file
MusicIndex = SoundIndex
End Sub

Public Sub StopMusic()
Call SoundStop(MusicIndex)
End Sub

