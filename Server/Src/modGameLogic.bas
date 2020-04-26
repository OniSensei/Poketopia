Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
On Error Resume Next
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Function FindOpenPlayerSlot() As Long
On Error Resume Next
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS

        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If

    Next

End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
On Error Resume Next
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
On Error Resume Next
    Dim i As Long
    TotalOnlinePlayers = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    Dim i As Long

    ' Check for subscript out of range
    If itemNum < 1 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(i, itemNum, ItemVal, mapnum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemNum < 0 Or itemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If itemNum >= 0 Then
            If itemNum <= MAX_ITEMS Then
                MapItem(mapnum, i).Num = itemNum
                MapItem(mapnum, i).value = ItemVal
                MapItem(mapnum, i).x = x
                MapItem(mapnum, i).y = y
                Set buffer = New clsBuffer
                buffer.WriteInteger SSpawnItem
                buffer.WriteLong i
                buffer.WriteLong itemNum
                buffer.WriteLong ItemVal
                buffer.WriteLong x
                buffer.WriteLong y
                SendDataToMap mapnum, buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
On Error Resume Next
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If item(map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(map(mapnum).Tile(x, y).Data1, map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Integer, ByVal High As Integer) As Integer
On Error Resume Next
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim NpcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    NpcNum = map(mapnum).NPC(MapNpcNum)

    If NpcNum > 0 Then
    
        MapNpc(mapnum).NPC(MapNpcNum).Num = NpcNum
        MapNpc(mapnum).NPC(MapNpcNum).Target = 0
        MapNpc(mapnum).NPC(MapNpcNum).TargetType = 0 ' clear
        
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) = GetNpcMaxVital(NpcNum, Vitals.hp)
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.mp) = GetNpcMaxVital(NpcNum, Vitals.mp)
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.SP) = GetNpcMaxVital(NpcNum, Vitals.SP)
        
        MapNpc(mapnum).NPC(MapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To map(mapnum).MaxX
            For y = 0 To map(mapnum).MaxY
                If map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If map(mapnum).Tile(x, y).Data1 = MapNpcNum Then
                        MapNpc(mapnum).NPC(MapNpcNum).x = x
                        MapNpc(mapnum).NPC(MapNpcNum).y = y
                        MapNpc(mapnum).NPC(MapNpcNum).Dir = map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, map(mapnum).MaxX)
                y = Random(0, map(mapnum).MaxY)
    
                If x > map(mapnum).MaxX Then x = map(mapnum).MaxX
                If y > map(mapnum).MaxY Then y = map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).NPC(MapNpcNum).x = x
                    MapNpc(mapnum).NPC(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To map(mapnum).MaxX
                For y = 0 To map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).NPC(MapNpcNum).x = x
                        MapNpc(mapnum).NPC(MapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set buffer = New clsBuffer
            buffer.WriteInteger SSpawnNpc
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Num
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        End If
        
        SendMapNpcVitals mapnum, MapNpcNum
    End If

End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
On Error Resume Next
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To MAX_PLAYERS

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).NPC(LoopI).Num > 0 Then
            If MapNpc(mapnum).NPC(LoopI).x = x Then
                If MapNpc(mapnum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, mapnum)
    Next

End Sub

Sub SpawnAllMapNpcs()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
On Error Resume Next
Exit Function
    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(Attacker).AttackTimer + item(GetPlayerEquipment(Attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(Victim, Vitals.hp) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack any player for thou art an admin!", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > ADMIN_MONITOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < 10 Then
        Call PlayerMsg(Attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < 10 Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanAttackPlayer = False
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean) As Boolean
On Error Resume Next
    Dim mapnum As Long
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim attackspeed As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(mapnum).NPC(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    'If MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.Hp) <= 0 Then
        'Exit Function
    'End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then

        ' attack speed from weapon
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            attackspeed = item(GetPlayerEquipment(Attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If NpcNum > 0 And GetTickCount > TempPlayer(Attacker).AttackTimer + attackspeed Then

            ' exit out early
            If IsSpell Then
                If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanAttackNpc = False
                    Exit Function
                End If
            End If
            
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y + 1
                Case DIR_DOWN
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y - 1
                Case DIR_LEFT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x + 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
                Case DIR_RIGHT
                    NpcX = MapNpc(mapnum).NPC(MapNpcNum).x - 1
                    NpcY = MapNpc(mapnum).NPC(MapNpcNum).y
            End Select

            If NpcX = GetPlayerX(Attacker) Then
                If NpcY = GetPlayerY(Attacker) Then
                    If NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NpcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        CanAttackNpc = False
                    Else
                        If Len(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                            'PlayerMsg Attacker, Trim$(NPC(NpcNum).Name) & ": " & Trim$(NPC(NpcNum).AttackSay), White
                            'MsgBox ReadText("Data\NPCScripts\" & player(Attacker).map & "I" & MapNpcNum & ".txt")
                            Dim a As String
                            a = ReadText("Data\NPCScripts\" & player(Attacker).map & "I" & MapNpcNum & ".txt")
                            If a = "" Then
                            Else
                            DoNpcScript Attacker, a
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
On Error Resume Next
    Dim mapnum As Long
    Dim NpcNum As Long

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    NpcNum = MapNpc(mapnum).NPC(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).NPC(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapnum).NPC(MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                CanNpcAttackPlayer = False
            Else

                If (GetPlayerY(index) - 1 = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                    CanNpcAttackPlayer = False
                Else

                    If (GetPlayerY(index) = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                        CanNpcAttackPlayer = False
                    Else

                        If (GetPlayerY(index) = MapNpc(mapnum).NPC(MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(mapnum).NPC(MapNpcNum).x) Then
                            CanNpcAttackPlayer = False
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

Function CanNpcAttackNpc(ByVal mapnum As Long, ByVal Attacker As Long, ByVal Victim As Long) As Boolean
On Error Resume Next
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long
    
    CanNpcAttackNpc = False

    ' Check for subscript out of range
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then
        Exit Function
    End If
    
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(mapnum).NPC(Attacker).Num <= 0 Then
        Exit Function
    End If
    
    ' Check for subscript out of range
    If MapNpc(mapnum).NPC(Victim).Num <= 0 Then
        Exit Function
    End If

    aNpcNum = MapNpc(mapnum).NPC(Attacker).Num
    vNpcNum = MapNpc(mapnum).NPC(Victim).Num
    
    If aNpcNum <= 0 Then Exit Function
    If vNpcNum <= 0 Then Exit Function

    ' Make sure the npcs arent already dead
    If MapNpc(mapnum).NPC(Attacker).Vital(Vitals.hp) <= 0 Then
        Exit Function
    End If
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(Victim).Vital(Vitals.hp) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).NPC(Attacker).AttackTimer + 1000 Then
        Exit Function
    End If
    
    MapNpc(mapnum).NPC(Attacker).AttackTimer = GetTickCount
    
    AttackerX = MapNpc(mapnum).NPC(Attacker).x
    AttackerY = MapNpc(mapnum).NPC(Attacker).y
    VictimX = MapNpc(mapnum).NPC(Victim).x
    VictimY = MapNpc(mapnum).NPC(Victim).y

    ' Check if at same coordinates
    If (VictimY + 1 = AttackerY) And (VictimX = AttackerX) Then
        CanNpcAttackNpc = True
    Else

        If (VictimY - 1 = AttackerY) And (VictimX = AttackerX) Then
            CanNpcAttackNpc = True
        Else

            If (VictimY = AttackerY) And (VictimX + 1 = AttackerX) Then
                CanNpcAttackNpc = True
            Else

                If (VictimY = AttackerY) And (VictimX - 1 = AttackerX) Then
                    CanNpcAttackNpc = True
                End If
            End If
        End If
    End If

End Function

Sub NpcAttackNpc(ByVal mapnum As Long, ByVal Attacker As Long, ByVal Victim As Long, ByVal damage As Long)
On Error Resume Next
    Dim i As Long
    Dim buffer As clsBuffer
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim n As Long
    
    If Attacker <= 0 Or Attacker > MAX_MAP_NPCS Then Exit Sub
    If Victim <= 0 Or Victim > MAX_MAP_NPCS Then Exit Sub
    
    If damage <= 0 Then Exit Sub
    
    aNpcNum = MapNpc(mapnum).NPC(Attacker).Num
    vNpcNum = MapNpc(mapnum).NPC(Victim).Num
    
    If aNpcNum <= 0 Then Exit Sub
    If vNpcNum <= 0 Then Exit Sub
    
    ' Send this packet so they can see the person attacking
    Set buffer = New clsBuffer
    buffer.WriteInteger SNpcAttack
    buffer.WriteLong Attacker
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing

    If damage >= MapNpc(mapnum).NPC(Victim).Vital(Vitals.hp) Then
        SendActionMsg mapnum, "-" & damage, BrightRed, 1, (MapNpc(mapnum).NPC(Victim).x * 32), (MapNpc(mapnum).NPC(Victim).y * 32)
        
        ' npc is dead.
        'Call GlobalMsg(CheckGrammar(Trim$(Npc(vNpcNum).Name), 1) & " has been killed by " & CheckGrammar(Trim$(Npc(aNpcNum).Name)) & "!", BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(Attacker).Target = 0
        MapNpc(mapnum).NPC(Attacker).TargetType = 0
        
        ' Drop the goods if they get it
        n = Int(Rnd * NPC(vNpcNum).DropChance) + 1
        If n = 1 Then
            Call SpawnItem(NPC(vNpcNum).DropItem, NPC(vNpcNum).DropItemValue, mapnum, MapNpc(mapnum).NPC(Victim).x, MapNpc(mapnum).NPC(Victim).y)
        End If
        
        ' Reset victim's stuff so it dies in loop
        MapNpc(mapnum).NPC(Victim).Num = 0
        MapNpc(mapnum).NPC(Victim).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(Victim).Vital(Vitals.hp) = 0
        
        ' send npc death packet to map
        Set buffer = New clsBuffer
        buffer.WriteInteger SNpcDead
        buffer.WriteLong Victim
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
    Else
        ' npc not dead, just do the damage
        MapNpc(mapnum).NPC(Victim).Vital(Vitals.hp) = MapNpc(mapnum).NPC(Victim).Vital(Vitals.hp) - damage
        ' Say damage
        SendActionMsg mapnum, "-" & damage, BrightRed, 1, (MapNpc(mapnum).NPC(Victim).x * 32), (MapNpc(mapnum).NPC(Victim).y * 32)
    End If

End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal damage As Long)
On Error Resume Next
    Dim Name As String
    Dim EXP As Long
    Dim mapnum As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(MapNpcNum).Num).Name)
    ' Send this packet so they can see the person attacking
    Set buffer = New clsBuffer
    buffer.WriteInteger SNpcAttack
    buffer.WriteLong MapNpcNum
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
    
    If damage <= 0 Then
        SendActionMsg GetPlayerMap(Victim), "BLOCK!", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        Exit Sub
    End If

    If damage >= GetPlayerVital(Victim, Vitals.hp) Then
        ' Say damage
        'Call PlayerMsg(Victim, CheckGrammar(Name, 1) & " hit you for " & Damage & " hit points.", BrightRed)
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & CheckGrammar(Name), BrightRed)
        ' Calculate exp to give attacker
        EXP = GetPlayerExp(Victim) \ 3

        ' Make sure we dont get less then 0
        If EXP < 0 Then EXP = 0
        If EXP = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & EXP & " exp.", BrightRed)
        End If

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(MapNpcNum).Target = 0
        MapNpc(mapnum).NPC(MapNpcNum).TargetType = 0
        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.hp, GetPlayerVital(Victim, Vitals.hp) - damage)
        Call SendVital(Victim, Vitals.hp)
        Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
        ' Say damage
        'Call PlayerMsg(Victim, CheckGrammar(Name, 1) & " hit you for " & Damage & " hit points.", BrightRed)
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    
    

    x = MapNpc(mapnum).NPC(MapNpcNum).x
    y = MapNpc(mapnum).NPC(MapNpcNum).y
    CanNpcMove = True
     
    If NPC(MapNpc(mapnum).NPC(MapNpcNum).Num).CanMove = NO Then Exit Function
     
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS

                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If

                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).Num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If

                Next

            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < map(mapnum).MaxY Then
                n = map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS

                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If

                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).Num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If

                Next

            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS

                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If

                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).Num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x - 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If

                Next

            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < map(mapnum).MaxX Then
                n = map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To MAX_PLAYERS

                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mapnum) And (GetPlayerX(i) = MapNpc(mapnum).NPC(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If

                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS

                    If (i <> MapNpcNum) And (MapNpc(mapnum).NPC(i).Num > 0) And (MapNpc(mapnum).NPC(i).x = MapNpc(mapnum).NPC(MapNpcNum).x + 1) And (MapNpc(mapnum).NPC(i).y = MapNpc(mapnum).NPC(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If

                Next

            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    
    If NPC(MapNpc(mapnum).NPC(MapNpcNum).Num).CanMove = NO Then Exit Sub
    
    MapNpc(mapnum).NPC(MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).NPC(MapNpcNum).y = MapNpc(mapnum).NPC(MapNpcNum).y - 1
            Set buffer = New clsBuffer
            buffer.WriteInteger SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong Movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).NPC(MapNpcNum).y = MapNpc(mapnum).NPC(MapNpcNum).y + 1
            Set buffer = New clsBuffer
            buffer.WriteInteger SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong Movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).NPC(MapNpcNum).x = MapNpc(mapnum).NPC(MapNpcNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteInteger SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong Movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).NPC(MapNpcNum).x = MapNpc(mapnum).NPC(MapNpcNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteInteger SNpcMove
            buffer.WriteLong MapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Dir
            buffer.WriteLong Movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
    End Select

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(MapNpcNum).Dir = Dir
    Set buffer = New clsBuffer
    buffer.WriteInteger SNpcDir
    buffer.WriteLong MapNpcNum
    buffer.WriteLong Dir
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
On Error Resume Next
    Dim i As Long
    Dim n As Long
    n = 0

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    Dim x As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case hp
            GetNpcMaxVital = NPC(NpcNum).hp
        Case mp
            GetNpcMaxVital = NPC(NpcNum).Stat(Stats.intelligence) * 2
        Case SP
            GetNpcMaxVital = NPC(NpcNum).Stat(Stats.spirit) * 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case hp
            i = NPC(NpcNum).Stat(Stats.vitality) \ 3

            If i < 1 Then i = 1
            GetNpcVitalRegen = i
    End Select

End Function

Sub ClearTempTiles()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        ClearTempTile i
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
On Error Resume Next
    Dim y As Long
    Dim x As Long
    TempTile(mapnum).DoorTimer = 0
    ReDim TempTile(mapnum).DoorOpen(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)

    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY
            TempTile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
On Error Resume Next
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY

            If map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Public Sub NewDoEvents()
On Error Resume Next
    'If GetQueueStatus(nLng) <> 0 Then DoEvents
    DoEvents
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal OldSlot As Integer, ByVal NewSlot As Integer)
On Error Resume Next
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If OldSlot = 0 Or NewSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, OldSlot)
    OldValue = GetPlayerInvItemValue(index, OldSlot)
    NewNum = GetPlayerInvItemNum(index, NewSlot)
    NewValue = GetPlayerInvItemValue(index, NewSlot)
    SetPlayerInvItemNum index, NewSlot, OldNum
    SetPlayerInvItemValue index, NewSlot, OldValue
    SetPlayerInvItemNum index, OldSlot, NewNum
    SetPlayerInvItemValue index, OldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)
On Error Resume Next
    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveItem index, GetPlayerEquipment(index, EqSlot), 1
        PlayerMsg index, "You unequip " & CheckGrammar(item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If
SendInventory index
End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
On Error Resume Next
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Function isInRange(ByVal range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
On Error Resume Next
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= range Then isInRange = True
End Function
