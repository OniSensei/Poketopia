Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
On Error Resume Next
    Dim i As Long
    Dim Tick As Long
    Dim tmr25 As Long
    Dim tmr500 As Long
    Dim tmr1000 As Long
    Dim LastUpdateSavePlayers As Long
    Dim LastUpdateMapSpawnItems As Long
    Dim LastUpdatePlayerVitals As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        
        If Tick > tmr25 Then
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).SpellBuffer > 0 Then
                        If GetTickCount > TempPlayer(i).SpellBufferTimer + (Spell(player(i).Spell(TempPlayer(i).SpellBuffer)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).SpellBuffer
                            TempPlayer(i).SpellBuffer = 0
                            TempPlayer(i).SpellBufferTimer = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                End If
            Next
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then

            For i = 1 To MAX_PLAYERS

                If frmServer.Socket(i).state > sckConnected Then
                    Call CloseSocket(i)
                End If

            Next
            
            UpdateNpcAI
            tmr500 = GetTickCount + 500
        End If
        'Fishing
        For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
        If Tick > TempPlayer(i).FishingTimer Then
        If TempPlayer(i).CanFish = False Then
        TempPlayer(i).FishingTimer = GetTickCount + 1000
        TempPlayer(i).CanFish = True
        End If
        End If
        End If
        Next

        If Tick > tmr1000 Then
        
            If isShuttingDown Then
                Call HandleShutdown
            End If

            tmr1000 = GetTickCount + 1000
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 1 minute - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
            AddMinutePlaytime i
            End If
            Next
            LastUpdateSavePlayers = GetTickCount + 60000
        End If

        Sleep 1
        NewDoEvents
    Loop

End Sub

Private Sub UpdateMapSpawnItems()
On Error Resume Next
    Dim x As Long
    Dim y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next

            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If

        NewDoEvents
    Next

End Sub

Private Sub UpdateNpcAI()
On Error Resume Next
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim n As Long
    Dim x1 As Long
    Dim y1 As Long
    Dim TickCount As Long
    Dim damage As Long
    Dim DistanceX As Long
    Dim DistanceY As Long
    Dim NpcNum As Long
    Dim Target As Long
    Dim TargetType As Byte
    Dim DidWalk As Boolean
    Dim buffer As clsBuffer
    Dim Resource_index As Long
    Dim TargetX As Long
    Dim TargetY As Long
    Dim target_verify As Boolean

    For y = 1 To MAX_MAPS
    
        '  Close the doors
        If TickCount > TempTile(y).DoorTimer + 5000 Then

            For x1 = 0 To map(y).MaxX
                For y1 = 0 To map(y).MaxY
                    If map(y).Tile(x1, y1).Type = TILE_TYPE_KEY And TempTile(y).DoorOpen(x1, y1) = YES Then
                        TempTile(y).DoorOpen(x1, y1) = NO
                        Set buffer = New clsBuffer
                        buffer.WriteInteger SMapKey
                        buffer.WriteLong x1
                        buffer.WriteLong y1
                        buffer.WriteLong 0
                        SendDataToMap y, buffer.ToArray()
                        Set buffer = Nothing
                    End If
                Next
            Next

        End If

        ' Respawning Resources
        If ResourceCache(y).Resource_Count > 0 Then
            For i = 0 To ResourceCache(y).Resource_Count
                Resource_index = map(y).Tile(ResourceCache(y).ResourceData(i).x, ResourceCache(y).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(y).ResourceData(i).ResourceState = 1 Or ResourceCache(y).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(y).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(y).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(y).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(y).ResourceData(i).cur_health = Resource(Resource_index).health
                            SendResourceCacheToMap y, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(y) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                NpcNum = MapNpc(y).NPC(x).Num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).NPC(x) > 0 And MapNpc(y).NPC(x).Num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then

                        For i = 1 To MAX_PLAYERS
                            If IsPlaying(i) Then
                                If GetPlayerMap(i) = y And MapNpc(y).NPC(x).Target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                    n = NPC(NpcNum).range
                                    DistanceX = MapNpc(y).NPC(x).x - GetPlayerX(i)
                                    DistanceY = MapNpc(y).NPC(x).y - GetPlayerY(i)

                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1

                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= n And DistanceY <= n Then
                                        If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                            If LenB(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                                                Call PlayerMsg(i, CheckGrammar(Trim$(NPC(NpcNum).Name), 1) & " says, '" & Trim$(NPC(NpcNum).AttackSay) & "' to you.", SayColor)
                                            End If
                                            MapNpc(y).NPC(x).TargetType = 1 ' player
                                            MapNpc(y).NPC(x).Target = i
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        
                        ' Check if target was found for NPC targetting
                        If MapNpc(y).NPC(x).Target = 0 Then
                            ' make sure it belongs to a faction
                            If NPC(NpcNum).faction > 0 Then
                                ' search for npc of another faction to target
                                For i = 1 To MAX_MAP_NPCS
                                    ' exist?
                                    If MapNpc(y).NPC(i).Num > 0 Then
                                        ' different faction?
                                        If NPC(MapNpc(y).NPC(i).Num).faction > 0 Then
                                            If NPC(MapNpc(y).NPC(i).Num).faction <> NPC(NpcNum).faction Then
                                                n = NPC(NpcNum).range
                                                DistanceX = MapNpc(y).NPC(x).x - CLng(MapNpc(y).NPC(i).x)
                                                DistanceY = MapNpc(y).NPC(x).y - CLng(MapNpc(y).NPC(i).y)
                                                
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                                
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n And DistanceY <= n Then
                                                    If NPC(NpcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Then
                                                        MapNpc(y).NPC(x).TargetType = 2 ' npc
                                                        MapNpc(y).NPC(x).Target = i
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).NPC(x) > 0 And MapNpc(y).NPC(x).Num > 0 Then
                    If MapNpc(y).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(y).NPC(x).StunTimer + (MapNpc(y).NPC(x).StunDuration * 1000) Then
                            MapNpc(y).NPC(x).StunDuration = 0
                            MapNpc(y).NPC(x).StunTimer = 0
                        End If
                    Else
                            
                        Target = MapNpc(y).NPC(x).Target
                        TargetType = MapNpc(y).NPC(x).TargetType
    
                        ' Check to see if its time for the npc to walk
                        If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If TargetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If Target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If IsPlaying(Target) And GetPlayerMap(Target) = y Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(Target)
                                        TargetX = GetPlayerX(Target)
                                    Else
                                        MapNpc(y).NPC(x).TargetType = 0 ' clear
                                        MapNpc(y).NPC(x).Target = 0
                                    End If
                                End If
                            
                            ElseIf TargetType = 2 Then 'npc
                                
                                If Target > 0 Then
                                    
                                    If MapNpc(y).NPC(Target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(y).NPC(Target).y
                                        TargetX = MapNpc(y).NPC(Target).x
                                    Else
                                        MapNpc(y).NPC(x).TargetType = 0 ' clear
                                        MapNpc(y).NPC(x).Target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                
                                i = Int(Rnd * 5)
                                If NPC(MapNpc(y).NPC(x).Num).CanMove = YES Then
                                ' Lets move the npc
                                Select Case i
                                    Case 0
    
                                        ' Up
                                        If MapNpc(y).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(y).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(y).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(y).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 1
    
                                        ' Right
                                        If MapNpc(y).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(y).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(y).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(y).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 2
    
                                        ' Down
                                        If MapNpc(y).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(y).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(y).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Left
                                        If MapNpc(y).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                    Case 3
    
                                        ' Left
                                        If MapNpc(y).NPC(x).x > TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_LEFT) Then
                                                Call NpcMove(y, x, DIR_LEFT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Right
                                        If MapNpc(y).NPC(x).x < TargetX And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_RIGHT) Then
                                                Call NpcMove(y, x, DIR_RIGHT, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Up
                                        If MapNpc(y).NPC(x).y > TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_UP) Then
                                                Call NpcMove(y, x, DIR_UP, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                        ' Down
                                        If MapNpc(y).NPC(x).y < TargetY And Not DidWalk Then
                                            If CanNpcMove(y, x, DIR_DOWN) Then
                                                Call NpcMove(y, x, DIR_DOWN, MOVING_WALKING)
                                                DidWalk = True
                                            End If
                                        End If
    
                                End Select
                                End If
                                ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(y).NPC(x).x - 1 = TargetX And MapNpc(y).NPC(x).y = TargetY Then
                                        If MapNpc(y).NPC(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(y, x, DIR_LEFT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(y).NPC(x).x + 1 = TargetX And MapNpc(y).NPC(x).y = TargetY Then
                                        If MapNpc(y).NPC(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(y, x, DIR_RIGHT)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(y).NPC(x).x = TargetX And MapNpc(y).NPC(x).y - 1 = TargetY Then
                                        If MapNpc(y).NPC(x).Dir <> DIR_UP Then
                                            Call NpcDir(y, x, DIR_UP)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    If MapNpc(y).NPC(x).x = TargetX And MapNpc(y).NPC(x).y + 1 = TargetY Then
                                        If MapNpc(y).NPC(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(y, x, DIR_DOWN)
                                        End If
    
                                        DidWalk = True
                                    End If
    
                                    ' We could not move so Target must be behind something, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
    
                                        If i = 1 Then
                                            i = Int(Rnd * 4)
    
                                            If CanNpcMove(y, x, i) Then
                                                Call NpcMove(y, x, i, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                    
                                End If
    
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(y, x, i) Then
                                        Call NpcMove(y, x, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If map(y).NPC(x) > 0 And MapNpc(y).NPC(x).Num > 0 Then
                    Target = MapNpc(y).NPC(x).Target
                    TargetType = MapNpc(y).NPC(x).TargetType

                    ' Check if the npc can attack the targeted player player
                    If Target > 0 Then
                    
                        If TargetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = y Then
    
                                ' Can the npc attack the player?
                                If CanNpcAttackPlayer(x, Target) Then
                                    If Not CanPlayerBlockHit(Target) Then
                                        damage = NPC(NpcNum).Stat(Stats.strength) - GetPlayerProtection(Target)
                                        Call NpcAttackPlayer(x, Target, damage)
                                    Else
                                        'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerEquipment(Target, Shield)).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                        SendActionMsg GetPlayerMap(Target), "BLOCK!", Cyan, 1, (GetPlayerX(Target) * 32), (GetPlayerY(Target) * 32)
                                    End If
                                End If
    
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(y).NPC(x).Target = 0
                                MapNpc(y).NPC(x).TargetType = 0 ' clear
                            End If
                        Else
                            If MapNpc(y).NPC(Target).Num > 0 Then ' npc exists
    
                                ' Can the npc attack the npc?
                                If CanNpcAttackNpc(y, x, Target) Then
                                        damage = CLng(NPC(NpcNum).Stat(Stats.strength)) - CLng(NPC(Target).Stat(Stats.endurance))
                                        If damage < 1 Then damage = 1
                                        Call NpcAttackNpc(y, x, Target, damage)
                                End If
    
                            Else
                                ' npc is dead or non-existant
                                MapNpc(y).NPC(x).Target = 0
                                MapNpc(y).NPC(x).TargetType = 0 ' clear
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNpc(y).NPC(x).Num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNpc(y).NPC(x).Vital(Vitals.hp) > 0 Then
                        MapNpc(y).NPC(x).Vital(Vitals.hp) = MapNpc(y).NPC(x).Vital(Vitals.hp) + GetNpcVitalRegen(NpcNum, Vitals.hp)

                        ' Check if they have more then they should and if so just set it to max
                        If MapNpc(y).NPC(x).Vital(Vitals.hp) > GetNpcMaxVital(NpcNum, Vitals.hp) Then
                            MapNpc(y).NPC(x).Vital(Vitals.hp) = GetNpcMaxVital(NpcNum, Vitals.hp)
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(y).NPC(x).Num = 0 And map(y).NPC(x) > 0 Then
                    If TickCount > MapNpc(y).NPC(x).SpawnWait + (NPC(map(y).NPC(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, y)
                    End If
                End If

            Next

        End If

        NewDoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub

Private Sub UpdatePlayerVitals()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerVital(i, Vitals.hp) <> GetPlayerMaxVital(i, Vitals.hp) Then
                Call SetPlayerVital(i, Vitals.hp, GetPlayerVital(i, Vitals.hp) + GetPlayerVitalRegen(i, Vitals.hp))
                Call SendVital(i, Vitals.hp)
            End If

            If GetPlayerVital(i, Vitals.mp) <> GetPlayerMaxVital(i, Vitals.mp) Then
                Call SetPlayerVital(i, Vitals.mp, GetPlayerVital(i, Vitals.mp) + GetPlayerVitalRegen(i, Vitals.mp))
                Call SendVital(i, Vitals.mp)
            End If

            If GetPlayerVital(i, Vitals.SP) <> GetPlayerMaxVital(i, Vitals.SP) Then
                Call SetPlayerVital(i, Vitals.SP, GetPlayerVital(i, Vitals.SP) + GetPlayerVitalRegen(i, Vitals.SP))
                Call SendVital(i, Vitals.SP)
            End If
        End If

    Next

End Sub

Private Sub UpdateSavePlayers()
On Error Resume Next
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Call TextAdd("Saving all online players...")
        'Call GlobalMsg("Saving all online players...", BrightBlue)

        For i = 1 To MAX_PLAYERS

            If IsPlaying(i) Then
                Call SavePlayer(i)
                If TempPlayer(i).eggExpTemp >= 1000 Then
             SaveEggFromTemp i
            End If
            If TempPlayer(i).eggStepsTemp >= 100 Then
            SaveEggFromTemp i
            End If
            End If
            
            

            NewDoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()
On Error Resume Next
    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If

    Secs = Secs - 1

    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

End Sub
