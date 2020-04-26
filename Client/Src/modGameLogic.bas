Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
    Dim FrameTime As Long
    Dim tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long, a As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long
    Dim tmr100 As Long
    Dim tmr120 As Long
    Dim tmr10000 As Long
    Dim tmrweather As Long
    ' *** Start GameLoop ***
    Do While InGame
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then

            ' characters
            If NumCharacters > 0 Then
                For i = 1 To NumCharacters    'Check to unload surfaces
                    If CharacterTimer(i) > 0 Then 'Only update surfaces in use
                        If CharacterTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Character(i)), LenB(DDSD_Character(i)))
                            Set DDS_Character(i) = Nothing
                            CharacterTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' Paperdolls
            If NumPaperdolls > 0 Then
                For i = 1 To NumPaperdolls    'Check to unload surfaces
                    If PaperdollTimer(i) > 0 Then 'Only update surfaces in use
                        If PaperdollTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Paperdoll(i)), LenB(DDSD_Paperdoll(i)))
                            Set DDS_Paperdoll(i) = Nothing
                            PaperdollTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            'Battle info
            
            
            
            ' animations
            If NumAnimations > 0 Then
                For i = 1 To NumAnimations    'Check to unload surfaces
                    If AnimationTimer(i) > 0 Then 'Only update surfaces in use
                        If AnimationTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Animation(i)), LenB(DDSD_Animation(i)))
                            Set DDS_Animation(i) = Nothing
                            AnimationTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Items
            If NumItems > 0 Then
                For i = 1 To NumItems    'Check to unload surfaces
                    If ItemTimer(i) > 0 Then 'Only update surfaces in use
                        If ItemTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Item(i)), LenB(DDSD_Item(i)))
                            Set DDS_Item(i) = Nothing
                            ItemTimer(i) = 0
                        End If
                    End If
                Next
            End If

            ' Resources
            If NumResources > 0 Then
                For i = 1 To NumResources    'Check to unload surfaces
                    If ResourceTimer(i) > 0 Then 'Only update surfaces in use
                        If ResourceTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_Resource(i)), LenB(DDSD_Resource(i)))
                            Set DDS_Resource(i) = Nothing
                            ResourceTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' spell icons
            If NumSpellIcons > 0 Then
                For i = 1 To NumSpellIcons    'Check to unload surfaces
                    If SpellIconTimer(i) > 0 Then 'Only update surfaces in use
                        If SpellIconTimer(i) < tick Then   'Unload the surface
                            Call ZeroMemory(ByVal VarPtr(DDSD_SpellIcon(i)), LenB(DDSD_SpellIcon(i)))
                            Set DDS_SpellIcon(i) = Nothing
                            SpellIconTimer(i) = 0
                        End If
                    End If
                Next
            End If
            
            ' door
            If DoorTimer > 0 Then
                If DoorTimer < tick Then
                    Call ZeroMemory(ByVal VarPtr(DDSD_Door), LenB(DDSD_Door))
                    Set DDS_Door = Nothing
                    DoorTimer = 0
                End If
            End If
            
            ' blood
            If BloodTimer > 0 Then
                If BloodTimer < tick Then
                    Call ZeroMemory(ByVal VarPtr(DDSD_Blood), LenB(DDSD_Blood))
                    Set DDS_Blood = Nothing
                    BloodTimer = 0
                End If
            End If

            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMainGame.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                                BltPlayerSpells
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = tick + 250
            End If
            
            ' Update inv animation
            If NumItems > 0 Then
                If tmr100 < tick Then
                    BltAnimatedInvItems
                    tmr100 = tick + 100
                End If
            End If
            
            For i = 1 To MAX_BYTE
                CheckAnimInstance i
            Next
            If ReceivingTime > 0 Then
            If ReceivingTime < tick Then
  
            End If
            End If
            tmr25 = tick + 25
        End If

        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then

            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    Call ProcessMovement(i)
                End If
            Next i

            ' Process npc movements (actually move them)
            For i = 1 To MAX_MAP_NPCS
                If map.NPC(i) > 0 Then
                    Call ProcessNpcMovement(i)
                End If
            Next i

            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If
        'Check weather
        
        If tmr120 < tick Then
        If inBattle = True Then
        If EnemyPokeImgLoaded = True And PokeImgLoaded = True Then
        
        EnemyPokeImg.ImageIndex = EnemyPokeImg.ImageIndex + 1
        If EnemyPokeImg.ImageIndex = EnemyPokeImg.ImageCount Then
        EnemyPokeImg.ImageIndex = 1
        End If
        
        PokeImg.ImageIndex = PokeImg.ImageIndex + 1
        If PokeImg.ImageIndex = PokeImg.ImageCount Then
        PokeImg.ImageIndex = 1
        End If
        
        
        Else
        If EnemyPokeImg.ImageIndex <> 1 Then
        EnemyPokeImg.ImageIndex = 1
        End If
        If PokeImg.ImageIndex <> 1 Then
        PokeImg.ImageIndex = 1
        End If
        End If
        tmr120 = tick + 80
        End If
        End If
        
        If tmrweather < tick Then
        For i = 1 To MAX_MAPS
        If Weather(i).Pics > 0 Then
        For a = 1 To Weather(i).Pics
        If Weather(i).pics_Y(a) >= 640 Then
        Weather(i).pics_Y(a) = 0
        Weather(i).pics_x(a) = Rand(16, 700)
        End If
         Weather(i).pics_Y(a) = Weather(i).pics_Y(a) + Weather(i).speed
        Next
        End If
        Next
        tmrweather = tick + 200
        
        End If
        
        'Check map music
        'GoranPlay(MapMusic)
        
        
        If MapMusic = "" Then
        If PlayingMapMusic <> "" Then
        StopPlay
        PlayingMapMusic = ""
        End If
        Else
        
        If MapMusic = PlayingMapMusic And MapMusic <> "" Then
        
        If IsMusicOver = True And Options.repeatmusic = YES Then
        PlayMapMusic MapMusic
        End If
        
        Else
        If MapMusic <> "" Then
        If Options.PlayMusic = YES Then
        PlayingMapMusic = MapMusic
        PlayMapMusic MapMusic
        End If
        End If
        End If
        
        End If
        
      
        

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < tick + 15
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If

    Loop

    frmMainGame.Visible = False
    frmChat.Visible = False
    If isLogging Then
        isLogging = False
        frmMainGame.picScreen.Visible = False
        frmMenu.Visible = True
        GettingMap = True
    Else
        ' Shutdown the game
        'frmSendGetData.Visible = True
         frmMainGame.lblSGInfo.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

End Sub

Sub ProcessMovement(ByVal Index As Long)
    Dim MovementSpeed As Long

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = RUN_SPEED '((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = WALK_SPEED '((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    If Player(Index).HasBike = YES Then MovementSpeed = BIKE_SPEED
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).YOffset = Player(Index).YOffset - MovementSpeed
            If Player(Index).YOffset < 0 Then Player(Index).YOffset = 0
        Case DIR_DOWN
            Player(Index).YOffset = Player(Index).YOffset + MovementSpeed
            If Player(Index).YOffset > 0 Then Player(Index).YOffset = 0
        Case DIR_LEFT
            Player(Index).XOffset = Player(Index).XOffset - MovementSpeed
            If Player(Index).XOffset < 0 Then Player(Index).XOffset = 0
        Case DIR_RIGHT
            Player(Index).XOffset = Player(Index).XOffset + MovementSpeed
            If Player(Index).XOffset > 0 Then Player(Index).XOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).XOffset >= 0) And (Player(Index).YOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        Else
            If (Player(Index).XOffset <= 0) And (Player(Index).YOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        End If
    End If

End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset < 0 Then MapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).YOffset > 0 Then MapNpc(MapNpcNum).YOffset = 0
                
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset < 0 Then MapNpc(MapNpcNum).XOffset = 0
                
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
                If MapNpc(MapNpcNum).XOffset > 0 Then MapNpc(MapNpcNum).XOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).dir = DIR_RIGHT Or MapNpc(MapNpcNum).dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).XOffset >= 0) And (MapNpc(MapNpcNum).YOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If MapNpc(MapNpcNum).Step = 0 Then
                        MapNpc(MapNpcNum).Step = 2
                    Else
                        MapNpc(MapNpcNum).Step = 0
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).XOffset <= 0) And (MapNpc(MapNpcNum).YOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
        
                    If MapNpc(MapNpcNum).Step = 0 Then
                        MapNpc(MapNpcNum).Step = 2
                    Else
                        MapNpc(MapNpcNum).Step = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub CheckMapGetItem()
    Dim Buffer As New clsBuffer
    Set Buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            Buffer.WriteLong CMapGetItem
            Buffer.WriteLong TCP_CODE
            SendData Buffer.ToArray()
        End If
    End If

    Set Buffer = Nothing
End Sub

Public Sub CheckAttack()
    Dim Buffer As clsBuffer
    Dim attackspeed As Long

    If ControlDown Then
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    '.Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                Set Buffer = New clsBuffer
                Buffer.WriteLong CAttack
                Buffer.WriteLong TCP_CODE
                SendData Buffer.ToArray()
                Set Buffer = Nothing
            End If
        End If
    End If

End Sub

Function IsTryingToMove() As Boolean

    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

End Function

Function CanMove() As Boolean
    Dim d As Long
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a battle.
    If BattleType > 0 Then
        CanMove = False
        Exit Function
    End If

    'Check if Storage or Bank is open
    If frmMainGame.picBank.Visible = True Or frmStorage.Visible = True Then
    CanMove = False
    Exit Function
    End If
    
    'Check if trainer card is opened
    If frmMainGame.picTrainerCard.Visible = True Then
    CanMove = False
    Exit Function
    End If
    
    If frmMainGame.picBattleCommands.Visible = True Then
    CanMove = False
    Exit Function
    End If
    
   
    d = GetPlayerDir(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If


End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
    Dim x As Long
    Dim y As Long
    Dim i As Long
    CheckDirection = False

    Select Case Direction
        Case DIR_UP
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            x = GetPlayerX(MyIndex)
            y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            x = GetPlayerX(MyIndex) - 1
            y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            x = GetPlayerX(MyIndex) + 1
            y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not
    If map.Tile(x, y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is tree or not
    If map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If map.Tile(x, y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(x, y).DoorOpen = NO Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
     ' Check to see if a npc is already on that tile
    For i = 1 To MAX_MAP_NPCS

        If MapNpc(i).num > 0 Then
            If MapNpc(i).x = x Then
                If MapNpc(i).y + 1 = y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Check to see if a player is already on that tile
    'For i = 1 To MAX_PLAYERS
        'If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            'If GetPlayerX(i) = X Then
                'If GetPlayerY(i) = Y Then
                    'CheckDirection = True
                    'Exit Function
                'End If
            'End If
       ' End If
   ' Next i

   

End Function

Sub CheckMovement()

    If IsTryingToMove Then
        If CanMove Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If
            
            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).YOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).XOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select
            
            If Player(MyIndex).XOffset = 0 Then
                If Player(MyIndex).YOffset = 0 Then
                    If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub PlayerSearch(ByVal CurX As Integer, ByVal CurY As Integer)
    Dim Buffer As clsBuffer

    If isInBounds Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CSearch
        Buffer.WriteLong TCP_CODE
        Buffer.WriteLong CurX
        Buffer.WriteLong CurY
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

Public Function isInBounds()

    If (CurX >= 0) Then
        If (CurX <= map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

End Function

Public Sub UpdateDrawMapName()
    
    DrawMapNameX = Camera.Left + ((MAX_MAPX + 1) * PIC_X / 2) - getWidth(TexthDC, Trim$(map.Name))
    DrawMapNameY = Camera.Top + 1

    Select Case map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = QBColor(BrightRed)
        Case MAP_MORAL_SAFE
            DrawMapNameColor = QBColor(White)
        Case Else
            DrawMapNameColor = QBColor(White)
    End Select

End Sub

Public Sub UseItem()

    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
End Sub

Public Sub ForgetSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellslot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellslot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellslot) > 0 Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong CForgetSpell
        Buffer.WriteLong TCP_CODE
        Buffer.WriteLong spellslot
        SendData Buffer.ToArray()
        Set Buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
End Sub

Public Sub CastSpell(ByVal spellslot As Long)
    Dim Buffer As clsBuffer

    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellslot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellslot)).MPCost Then
        Call AddText("Not enough MP to cast " & Trim$(Spell(PlayerSpells(spellslot)).Name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellslot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set Buffer = New clsBuffer
                Buffer.WriteLong CCast
                Buffer.WriteLong TCP_CODE
                Buffer.WriteLong spellslot
                SendData Buffer.ToArray()
                Set Buffer = Nothing
                SpellBuffer = spellslot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("Cannot cast while walking!", BrightRed)
            End If
        End If
    Else
        Call AddText("No spell here.", BrightRed)
    End If

End Sub

Sub ClearTempTile()
    Dim x As Long
    Dim y As Long
    ReDim TempTile(0 To map.MaxX, 0 To map.MaxY)

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            TempTile(x, y).DoorOpen = NO
        Next
    Next

End Sub

Public Sub DevMsg(ByVal text As String, ByVal color As Byte)

    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText("[DEV]" & text, color)
        End If
    End If

    Debug.Print text
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If

End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long

    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If

End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String

    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If

End Function

Sub DrawPing()
    Dim PingToDraw As String
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select

    
End Sub

Public Sub UpdateSpellWindow(ByVal spellnum As Long, ByVal x As Long, ByVal y As Long)

End Sub

Public Sub UpdateDescWindow(ByVal itemnum As Long, ByVal Amount As Long, ByVal x As Long, ByVal y As Long)

End Sub

Public Sub CacheResources()
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY

            If map.Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).x = x
                MapResource(Resource_Count).y = y
            End If

        Next
    Next

    Resource_Index = Resource_Count
End Sub

Public Sub CreateActionMsg(ByVal message As String, ByVal color As Integer, ByVal MsgType As Byte, ByVal x As Long, ByVal y As Long)

    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .message = message
        .color = color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .x = x
        .y = y
        .Width = Len(message)
        .Height = 16
        .backColor = DarkGrey
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).y = ActionMsg(ActionMsgIndex).y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).x = ActionMsg(ActionMsgIndex).x + Rand(-8, 8)
    End If

End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
    ActionMsg(Index).message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).x = 0
    ActionMsg(Index).y = 0
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long
    Dim lockindex As Long
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).FrameIndex(Layer) = 0 Then AnimInstance(Index).FrameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).FrameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).FrameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).FrameIndex(Layer) = AnimInstance(Index).FrameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    'frmMainGame.lblShopName.Caption = Trim$(Shop(shopnum).Name)
    InShop = shopnum
    ShopAction = 0
    
    frmMainGame.OpenMenu (MENU_SHOP)
    'BltShop
    frmMainGame.ShoploadItems
End Sub

Public Sub UpdateBattle()
On Error Resume Next
If BattlePokemon < 1 Or enemyPokemon.PokemonNumber < 1 Then Exit Sub
Dim i As Long
Dim Wait As Long
Dim pokeMoves As Long
Wait = GetTickCount
'Round
frmBattle.lblRound.Caption = BattleRound
    ' enemy pokemon
    If enemyPokemon.PokemonNumber > 0 Then
        frmBattle.lblEnemyName.Caption = Trim$(Pokemon(enemyPokemon.PokemonNumber).Name)
        frmBattle.EnemyImg(0).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\" & enemyPokemon.PokemonNumber & ".png")
    End If
    frmBattle.lblEnemyLevel = "Lvl." & enemyPokemon.Level
    If enemyPokemon.HP <= 0 Then
        frmBattle.lblEnemyHP.Caption = "0/" & enemyPokemon.MaxHp
        frmBattle.picEnemyPoke.Inverted = True
        frmBattle.shapeHPEnemy.Width = 0
    Else
    frmBattle.lblEnemyHP.Caption = enemyPokemon.HP & "/" & enemyPokemon.MaxHp
    frmBattle.picEnemyPoke.Inverted = False
    Dim eSW As Long
    eSW = (enemyPokemon.HP / enemyPokemon.MaxHp) * 2415
    frmBattle.shapeHPEnemy.Width = eSW
    End If
    'your pokemon
    frmBattle.lblMyName.Caption = Trim$(Pokemon(PokemonInstance(BattlePokemon).PokemonNumber).Name)
    If PokemonInstance(BattlePokemon).HP < 0 Then
    frmBattle.lblMyHP.Caption = "0/" & PokemonInstance(BattlePokemon).MaxHp
    frmBattle.shapeHpMine.Width = 0
    Else
    frmBattle.lblMyHP.Caption = PokemonInstance(BattlePokemon).HP & "/" & PokemonInstance(BattlePokemon).MaxHp
    Dim mSW As Long
    mSW = (PokemonInstance(BattlePokemon).HP / PokemonInstance(BattlePokemon).MaxHp) * 2415
    frmBattle.shapeHpMine.Width = mSW
    End If
    
    frmBattle.lvlMyLevel.Caption = "Lvl." & PokemonInstance(BattlePokemon).Level
    For i = 1 To 4
    If PokemonInstance(BattlePokemon).moves(i).number < 1 Or PokemonInstance(BattlePokemon).moves(i).number > MAX_MOVES Then
    frmBattle.cmdMove(i).Caption = "None." & "(0)"
    frmMainGame.cmdPokeMove(i).Caption = "None." & "(0)"
    pokeMoves = pokeMoves + 1
    Else
    If PokemonInstance(BattlePokemon).moves(i).pp < 1 Then
    pokeMoves = pokeMoves + 1
    End If
    frmMainGame.cmdPokeMove(i).Caption = Trim$(PokemonMove(PokemonInstance(BattlePokemon).moves(i).number).Name) & " (" & PokemonInstance(BattlePokemon).moves(i).pp & ")"
    frmBattle.cmdMove(i).Caption = Trim$(PokemonMove(PokemonInstance(BattlePokemon).moves(i).number).Name) & " (" & PokemonInstance(BattlePokemon).moves(i).pp & ")"
    End If
    Next
    
    If pokeMoves = 4 Then 'This means pokemon doesn't have any PP or move,so we add struggle.
    pokeMoves = 0
       For i = 1 To 4
       frmBattle.cmdMove(i).Caption = "Struggle (Infinite)"
       frmMainGame.cmdPokeMove(i).Caption = "Struggle (Infinite)"
       Next
    End If
    
    
    For i = 1 To 6
    If PokemonInstance(i).PokemonNumber > 0 Then
    If PokemonInstance(i).HP > 0 Then
    frmBattle.imgSwitch(i).Picture = Nothing
    frmBattle.imgSwitch(i).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\" & PokemonInstance(i).PokemonNumber & ".png")
    frmBattle.imgSwitch(i).AnimateOnLoad = True
    Else
    frmBattle.imgSwitch(i).Picture = Nothing
    End If
    Else
    frmBattle.imgSwitch(i).Picture = Nothing
    End If
    Next
    
    'Load GUI
    frmMainGame.lblBattleEXP.Caption = Val(PokemonInstance(BattlePokemon).EXP) & "/" & Val(PokemonInstance(BattlePokemon).expNeeded)
    frmBattle.loadGUI PokemonInstance(BattlePokemon).PokemonNumber, enemyPokemon.PokemonNumber, PokemonInstance(BattlePokemon).HP, PokemonInstance(BattlePokemon).MaxHp, enemyPokemon.HP, enemyPokemon.MaxHp, enemyPokemon.isShiny, PokemonInstance(BattlePokemon).isShiny
    Call LoadBattleGDI
End Sub
Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then

            If Trim$(GetPlayerName(i)) = Trim$(Name) Then
            FindPlayer = i
            Exit Function
            End If
        End If

    Next

    FindPlayer = 0
End Function
