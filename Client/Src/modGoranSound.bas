Attribute VB_Name = "modGoranSound"
Dim PlayerIsPlaying As Boolean  'determine when the player is playing
Dim Player As FilgraphManager   'Reference to our player
Dim PlayerPos As IMediaPosition 'Reference to determine media position
Dim PlayerAU As IBasicAudio     'Reference to determine Audio Volume
Dim i As Integer                'Icon index
Sub GoranPlay(ByVal file As String)
If Options.PlayMusic = 0 Then Exit Sub 'Music isnt allowed :/
Dim CurState As Long
 'check player
 If Not Player Is Nothing Then
    'Get the state
    Player.GetState X, CurState
      
    If CurState = 1 Then
      PausePlay
      Exit Sub
    End If
 End If
 
 StartPlay (file) 'Start playing the file

End Sub

Sub PlayMapMusic(ByVal music As String)
StopPlay
If FileExist(App.Path & "\Data Files\music\" & music, True) And InMapEditor = False Then
GoranPlay (App.Path & "\Data Files\music\" & music)
End If
End Sub

Sub StartPlay(ByVal file As String)

On Error GoTo error                   'Handle Error
   'Set objects
   Set Player = New FilgraphManager   'Player
   Set PlayerPos = Player             'Position
   Set PlayerAU = Player              'Volume
   
   Player.RenderFile file   'Load file
   Player.Run                         'Run player
   
   PlayerIsPlaying = True 'We are playing
   If Not Player Is Nothing Then
    'if g_objMediaControl has been assigned

       'PlayerAU.Volume = GetVolume

    End If


Exit Sub
error:                                 'Handle error
StopPlay                            'Stop player
End Sub

Function GetPlayerDuration() As Long
On Error Resume Next
GetPlayerDuration = PlayerPos.Duration
End Function

Function GetPlayerPosition() As Long
On Error Resume Next
GetPlayerPosition = PlayerPos.CurrentPosition
End Function

Function IsMusicOver() As Boolean
If GetPlayerPosition >= GetPlayerDuration Then
IsMusicOver = True
End If
End Function

Sub StopPlay()

  If Player Is Nothing Then Exit Sub 'Not playing nothing to stop!
           'No timer after stop
  
  'Stop playing
  Player.Stop
  'Set time and status label
  
    
End Sub

Sub PausePlay()
   
Static Paused As Boolean                'If paused
Dim CurState As Long                    'Current state of the player

  If Player Is Nothing Then Exit Sub    'Not playing nothing to pause!
     
     'Get player state
     Player.GetState X, CurState
     
     If CurState = 2 Then
       'Is playing, pause it
       Paused = True
       Player.Pause
       
     Else
       'Is paused, run again
       Paused = False
       Player.Run
       
     End If
     
End Sub

