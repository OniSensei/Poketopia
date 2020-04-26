Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal Index As Long)
On Error Resume Next
    If Not IsPlaying(Index) Then
        Call JoinGame(Index)
        Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
        Call TextAdd(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & Options.Game_Name & ".")
        Call UpdateCaption
    End If
End Sub
Function HealPokemons(ByVal Index As Long) As Boolean
On Error Resume Next
Dim x As Long
Dim a As Long
Dim n As Long
Dim m As Long
Dim k As Long
Dim y As Long
Dim z As Long
Dim i As Long
Dim healPoke(1 To 6) As Boolean
For i = 1 To 6
healPoke(i) = False
Next



For i = 1 To 6
If player(Index).PokemonInstance(i).PokemonNumber > 0 Then
If player(Index).PokemonInstance(i).hp < player(Index).PokemonInstance(i).MaxHp Then
healPoke(i) = True
Else
For k = 1 To 4
If player(Index).PokemonInstance(i).moves(k).number > 0 Then
If player(Index).PokemonInstance(i).moves(k).pp <> PokemonMove(player(Index).PokemonInstance(i).moves(k).number).pp Then
healPoke(i) = True
End If
End If
Next
End If
End If
Next

a = 0
For i = 1 To 6
If healPoke(i) = True Then
a = a + 1
End If
Next

If a = 0 Then Exit Function



For x = 1 To 6
If player(Index).PokemonInstance(x).PokemonNumber > 0 Then
player(Index).PokemonInstance(x).hp = player(Index).PokemonInstance(x).MaxHp
                For m = 1 To 4
                If player(Index).PokemonInstance(x).moves(m).number > 0 Then
                player(Index).PokemonInstance(x).moves(m).pp = PokemonMove(player(Index).PokemonInstance(x).moves(m).number).pp
                End If
                player(Index).PokemonInstance(x).status = STATUS_NOTHING
                player(Index).PokemonInstance(x).turnsneed = 0
                player(Index).PokemonInstance(x).statusstun = 0
                Next
            End If
        Next
        If GetPlayerMap(Index) = 115 Then
        SendDialog Index, "Your pokemons are now healed!"
        Else
        SendDialog Index, "Nurse Joy: Your pokemons are now healed!"
        End If
        'PlayerMsg index, "Nurse Joy: Thank you", BrightGreen
        SendPlayerPokemon Index
        SendSound Index, "Heal.wav"
        HealPokemons = True
End Function















Sub AddTP(ByVal Index As Long, ByVal Stat As Long, ByVal pokeSlot As Long)
On Error Resume Next
'Dont add if there is no TP
If player(Index).PokemonInstance(pokeSlot).TP <= 0 Then
PlayerMsg Index, "Your pokemon doesn't have enough TP!", BrightRed
Exit Sub
End If
'If there is TP add stat

Select Case Stat
Case STAT_HP
player(Index).PokemonInstance(pokeSlot).MaxHp = player(Index).PokemonInstance(pokeSlot).MaxHp + 2
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
Case STAT_ATK
If player(Index).PokemonInstance(pokeSlot).atk >= player(Index).PokemonInstance(pokeSlot).MaxHp Then Exit Sub
player(Index).PokemonInstance(pokeSlot).atk = player(Index).PokemonInstance(pokeSlot).atk + 1
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
Case STAT_DEF
If player(Index).PokemonInstance(pokeSlot).def >= player(Index).PokemonInstance(pokeSlot).MaxHp Then Exit Sub
player(Index).PokemonInstance(pokeSlot).def = player(Index).PokemonInstance(pokeSlot).def + 1
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
Case STAT_SPATK
If player(Index).PokemonInstance(pokeSlot).spatk >= player(Index).PokemonInstance(pokeSlot).MaxHp Then Exit Sub
player(Index).PokemonInstance(pokeSlot).spatk = player(Index).PokemonInstance(pokeSlot).spatk + 1
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
Case STAT_SPDEF
If player(Index).PokemonInstance(pokeSlot).spdef >= player(Index).PokemonInstance(pokeSlot).MaxHp Then Exit Sub
player(Index).PokemonInstance(pokeSlot).spdef = player(Index).PokemonInstance(pokeSlot).spdef + 1
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
Case STAT_SPEED
If player(Index).PokemonInstance(pokeSlot).spd >= player(Index).PokemonInstance(pokeSlot).MaxHp Then Exit Sub
If player(Index).PokemonInstance(pokeSlot).spd >= player(Index).PokemonInstance(pokeSlot).def + player(Index).PokemonInstance(pokeSlot).spdef Then Exit Sub
player(Index).PokemonInstance(pokeSlot).spd = player(Index).PokemonInstance(pokeSlot).spd + 1
player(Index).PokemonInstance(pokeSlot).TP = player(Index).PokemonInstance(pokeSlot).TP - 1
End Select

SendPlayerPokemon (Index)


End Sub
Sub JoinGame(ByVal Index As Long)
On Error Resume Next
    Dim i As Long
    Dim buffer As clsBuffer
    ' Set the flag so we know the person is in the game
    TempPlayer(Index).InGame = True

    ' Send a global message that he/she joined
    If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
        'Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        'Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If

    'Update the log
    frmServer.lvwInfo.ListItems(Index).SubItems(1) = GetPlayerIP(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = GetPlayerLogin(Index)
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = GetPlayerName(Index)
    ' Send an ok to client to start receiving in game data
    Set buffer = New clsBuffer
    buffer.WriteInteger SLoginOk
    buffer.WriteLong Index
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
    TotalPlayersOnline = TotalPlayersOnline + 1
    ' Send some more little goodies, no need to explain these
    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendAnimations(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    'Call SendSpells(index)
    'Call SendResources(index)
    Call SendPokemon(Index)
    Call SendInventory(Index)
    Call SendWornEquipment(Index)
    Call SendMapEquipment(Index)
    Call SendPlayerPokemon(Index)
    Call SendMovesToPlayer(Index)
    Call SendPlayerStorage(Index)
    Call SendMusicToOne(Index)
    Call SendUpdateBank(Index)
    'Call SendNews(index)
    Call SendFlashlight(Index, GetMapFlashlight(player(Index).map))
   
    For i = 1 To Vitals.Vital_Count - 1
        'Call SendVital(index, i)
    Next
    
    SendEXP Index
    
    If TempPlayer(Index).HasBike = YES Then
    UseBike Index
    Else
    If TempPlayer(Index).HasBike <> YES And GetPlayerSprite(Index) = 514 Then
    SetPlayerSprite Index, 509
    End If
    End If
    
    


    'Call SendStats(index)
    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    ' Send welcome messages
 Call SendWelcome(Index)
 Call CheckPlayerMembership(Index)
 Call CheckForLoginItem(Index)
    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        'SendResourceCacheTo index, i
    Next
    ' Send the flag so they know they can start doing stuff
    Set buffer = New clsBuffer
    buffer.WriteInteger SInGame
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub LeftGame(ByVal Index As Long)
On Error Resume Next
    Dim n As Long
    Dim i As Long
    If TempPlayer(Index).InGame Then
        TempPlayer(Index).InGame = False

        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(Index)) < 1 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If

        ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If TempPlayer(Index).InParty = YES Then
            n = TempPlayer(Index).PartyPlayer
            'Call PlayerMsg(n, GetPlayerName(index) & " has left " & Options.Game_Name & ", disbanning party.", BrightBlue)
            TempPlayer(n).InParty = NO
            TempPlayer(n).PartyPlayer = 0
        End If
        
        If TempPlayer(Index).isTrading = YES Then
           For i = 1 To MAX_PLAYERS
             If IsPlaying(i) Then
                 If Trim$(TempPlayer(i).TradeName) = Trim$(player(Index).Name) Then
                    TempPlayer(i).isTrading = NO
                    TempPlayer(i).TradeName = ""
                    SendTradeStop i
                 End If
             End If
           Next
        End If
        
        If TempPlayer(Index).eggExpTemp >= 1000 Then
SaveEggFromTemp Index
End If
If TempPlayer(Index).eggStepsTemp >= 100 Then
    SaveEggFromTemp Index
   End If
   
   If TempPlayer(Index).HasBike = YES Then
   UseBike Index
   End If
        

        Call SavePlayer(Index)

        ' Send a global message that he/she left
        If GetPlayerAccess(Index) <= ADMIN_MONITOR Then
            'Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
        Else
            'Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
        End If

        Call TextAdd(GetPlayerName(Index) & " has disconnected from " & Options.Game_Name & ".")
        Call SendLeftGame(Index)
        TotalPlayersOnline = TotalPlayersOnline - 1
    End If

    Call ClearPlayer(Index)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal damage As Long, Optional ByVal spellnum As Long)
On Error Resume Next
    Dim Name As String
    Dim EXP As Long
    Dim n As Long
    Dim i As Long
    Dim str As Long
    Dim def As Long
    Dim mapnum As Long
    Dim NpcNum As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(mapnum).NPC(MapNpcNum).Num
    Name = Trim$(NPC(NpcNum).Name)
    
    If spellnum = 0 Then
        ' Send this packet so they can see the person attacking
        Set buffer = New clsBuffer
        buffer.WriteInteger SAttack
        buffer.WriteLong Attacker
        SendDataToMapBut Attacker, mapnum, buffer.ToArray()
        Set buffer = Nothing
    End If
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If

    If damage >= MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) Then
    
        SendActionMsg GetPlayerMap(Attacker), "-" & damage, BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        'SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y

        ' Calculate exp to give attacker
        EXP = NPC(NpcNum).EXP

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If

        ' Check if in party, if so divide the exp up by 2
        If TempPlayer(Attacker).InParty = NO Then
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            SendEXP Attacker
            'Call PlayerMsg(Attacker, "You have gained " & Exp & " experience points.", BrightBlue)
            SendActionMsg GetPlayerMap(Attacker), "+" & EXP & " EXP", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        Else
            EXP = EXP / 2

            If EXP < 0 Then
                EXP = 1
            End If

            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            SendEXP Attacker
            'Call PlayerMsg(Attacker, "You have gained " & Exp & " party experience points.", BrightBlue)
            SendActionMsg GetPlayerMap(Attacker), "+" & EXP & " Shared EXP", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            n = TempPlayer(Attacker).PartyPlayer

            If n > 0 Then
                Call SetPlayerExp(n, GetPlayerExp(n) + EXP)
                SendEXP n
                'Call PlayerMsg(n, "You have gained " & Exp & " party experience points.", BrightBlue)
                SendActionMsg GetPlayerMap(n), "+" & EXP & " EXP", White, 1, (GetPlayerX(n) * 32), (GetPlayerY(n) * 32)
            End If
        End If

        ' Drop the goods if they get it
        n = Int(Rnd * NPC(NpcNum).DropChance) + 1

        If n = 1 Then
            Call SpawnItem(NPC(NpcNum).DropItem, NPC(NpcNum).DropItemValue, mapnum, MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y)
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(MapNpcNum).Num = 0
        MapNpc(mapnum).NPC(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) = 0
        
        Set buffer = New clsBuffer
        buffer.WriteInteger SNpcDead
        buffer.WriteLong MapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
        
        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check for level up party member
        If TempPlayer(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(TempPlayer(Attacker).PartyPlayer)
        End If

        ' Check if target is npc that died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC Then
            If TempPlayer(Attacker).Target = MapNpcNum Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If

    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) = MapNpc(mapnum).NPC(MapNpcNum).Vital(Vitals.hp) - damage

        ' Check for a weapon and say damage
        SendActionMsg mapnum, "-" & damage, BrightRed, 1, (MapNpc(mapnum).NPC(MapNpcNum).x * 32), (MapNpc(mapnum).NPC(MapNpcNum).y * 32)
        'SendBlood GetPlayerMap(Attacker), MapNpc(mapnum).NPC(MapNpcNum).x, MapNpc(mapnum).NPC(MapNpcNum).y
        
        ' send animation
        If n > 0 Then
            If spellnum = 0 Then Call SendAnimation(mapnum, item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
        End If

        ' Check if we should send a message
        If MapNpc(mapnum).NPC(MapNpcNum).Target = 0 Then
            If LenB(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                Call PlayerMsg(Attacker, CheckGrammar(Trim$(NPC(NpcNum).Name), 1) & " says: " & Trim$(NPC(NpcNum).AttackSay), SayColor)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(mapnum).NPC(MapNpcNum).TargetType = 1 ' player
        MapNpc(mapnum).NPC(MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(MapNpcNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then

            For i = 1 To MAX_MAP_NPCS

                If MapNpc(mapnum).NPC(i).Num = MapNpc(mapnum).NPC(MapNpcNum).Num Then
                    MapNpc(mapnum).NPC(i).Target = Attacker
                    MapNpc(mapnum).NPC(i).TargetType = 1 ' player
                End If

            Next

        End If
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC MapNpcNum, mapnum, spellnum
        End If
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).AttackTimer = GetTickCount
    End If
End Sub

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal damage As Long, Optional ByVal spellnum As Long = 0)
On Error Resume Next
    Dim EXP As Long
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If

    ' Send this packet so they can see the person attacking
    Set buffer = New clsBuffer
    buffer.WriteInteger SAttack
    buffer.WriteLong Attacker
    SendDataToMapBut Attacker, GetPlayerMap(Attacker), buffer.ToArray()
    Set buffer = Nothing

    If damage >= GetPlayerVital(Victim, Vitals.hp) Then
    
        'Call PlayerMsg(Attacker, "You hit " & GetPlayerName(Victim) & " for " & Damage & " hit points.", White)
        'Call PlayerMsg(Victim, GetPlayerName(Attacker) & " hit you for " & Damage & " hit points.", BrightRed)
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker), BrightRed)
        ' Calculate exp to give attacker
        EXP = (GetPlayerExp(Victim) \ 10)

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If

        If EXP = 0 Then
            Call PlayerMsg(Victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(Attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
            SendEXP Victim
            Call PlayerMsg(Victim, "You lost " & EXP & " exp.", BrightRed)
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
            SendEXP Attacker
            Call PlayerMsg(Attacker, "You received " & EXP & " exp.", BrightBlue)
        End If

        ' Check for a level up
        Call CheckPlayerLevelUp(Attacker)

        ' Check if target is player who died and if so set target to 0
        If TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER Then
            If TempPlayer(Attacker).Target = Victim Then
                TempPlayer(Attacker).Target = 0
                TempPlayer(Attacker).TargetType = TARGET_TYPE_NONE
            End If
        End If

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerData(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(Victim)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.hp, GetPlayerVital(Victim, Vitals.hp) - damage)
        Call SendVital(Victim, Vitals.hp)
        SendActionMsg GetPlayerMap(Victim), "-" & damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer Victim, spellnum
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = GetTickCount
End Sub

Function GetPlayerDamage(ByVal Index As Long) As Long
On Error Resume Next
    Dim Weapon As Long
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    GetPlayerDamage = GetPlayerStat(Index, Stats.strength)

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        Weapon = GetPlayerEquipment(Index, Weapon)
        GetPlayerDamage = GetPlayerDamage + item(Weapon).Data2
    End If

End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
On Error Resume Next
    Dim Armor As Long
    Dim Helm As Long
    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    Armor = GetPlayerEquipment(Index, Armor)
    Helm = GetPlayerEquipment(Index, Helmet)
    GetPlayerProtection = (GetPlayerStat(Index, Stats.endurance) \ 5)

    If Armor > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Armor).Data2
    End If

    If Helm > 0 Then
        GetPlayerProtection = GetPlayerProtection + item(Helm).Data2
    End If

End Function

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
On Error Resume Next
    On Error Resume Next
    Dim i As Long
    Dim n As Long

    If GetPlayerEquipment(Index, Weapon) > 0 Then
        n = (Rnd) * 2

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.strength) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If

End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Dim ShieldSlot As Long
    ShieldSlot = GetPlayerEquipment(Index, Shield)

    If ShieldSlot > 0 Then
        n = Int(Rnd * 2)

        If n = 1 Then
            i = (GetPlayerStat(Index, Stats.endurance) \ 2) + (GetPlayerLevel(Index) \ 2)
            n = Int(Rnd * 100) + 1

            If n <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If

End Function

Public Sub BufferSpell(ByVal Index As Long, ByVal spellslot As Long)
On Error Resume Next
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim range As Long
    Dim HasBuffered As Boolean
    
    Dim TargetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(Index, spellslot)
    mapnum = GetPlayerMap(Index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(Index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.mp) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    TargetType = TempPlayer(Index).TargetType
    Target = TempPlayer(Index).Target
    range = Spell(spellnum).range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not Target > 0 Then
                PlayerMsg Index, "You do not have a target.", BrightRed
            End If
            If TargetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanAttackPlayer(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(range, GetPlayerX(Index), GetPlayerY(Index), MapNpc(mapnum).NPC(Target).x, MapNpc(mapnum).NPC(Target).y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanAttackNpc(Index, Target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation mapnum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        TempPlayer(Index).SpellBuffer = spellslot
        TempPlayer(Index).SpellBufferTimer = GetTickCount
        Exit Sub
    Else
        'SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal spellslot As Long)
On Error Resume Next
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim i As Long
    Dim AoE As Long
    Dim range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
    
    Dim TargetType As Byte
    Dim Target As Long
    
    Dim buffer As clsBuffer
    Dim SpellCastType As Long
    
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(Index, spellslot)
    mapnum = GetPlayerMap(Index)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then Exit Sub

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.mp) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = Spell(spellnum).Vital
    AoE = Spell(spellnum).AoE
    range = Spell(spellnum).range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.hp, True, Index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.mp, True, Index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(spellnum).map, Spell(spellnum).x, Spell(spellnum).y
                    SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(Index)
                y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                TargetType = TempPlayer(Index).TargetType
                Target = TempPlayer(Index).Target
    
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(Target)
                    y = GetPlayerY(Target)
                Else
                    x = MapNpc(mapnum).NPC(Target).x
                    y = MapNpc(mapnum).NPC(Target).y
                End If
                
                If Not isInRange(range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                    PlayerMsg Index, "Target not in range.", BrightRed
                    'SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                            If i <> Index Then
                                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanAttackPlayer(Index, i, True) Then
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            AttackPlayer Index, i, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(hp) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    If CanAttackNpc(Index, i, True) Then
                                        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        AttackNpc Index, i, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.hp
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.mp
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.mp
                        increment = False
                    End If
                    
                    DidCast = True
                    For i = 1 To MAX_PLAYERS
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(Index) Then
                                If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, increment, i, Vital, spellnum
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(i).Num > 0 Then
                            If MapNpc(mapnum).NPC(i).Vital(hp) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(i).x, MapNpc(mapnum).NPC(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
        
            TargetType = TempPlayer(Index).TargetType
            Target = TempPlayer(Index).Target

            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            
            If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(Target)
                y = GetPlayerY(Target)
            Else
                x = MapNpc(mapnum).NPC(Target).x
                y = MapNpc(mapnum).NPC(Target).y
            End If
                
            If Not isInRange(range, GetPlayerX(Index), GetPlayerY(Index), x, y) Then
                PlayerMsg Index, "Target not in range.", BrightRed
                'S'endClearSpellBuffer index
                Exit Sub
            End If
            
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER Then
                        If CanAttackPlayer(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                AttackPlayer Index, Target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanAttackNpc(Index, Target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                AttackNpc Index, Target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.mp
                        increment = False
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.mp
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.hp
                        increment = True
                    End If
                    
                    If TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanAttackPlayer(Index, Target, True) Then
                                SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, Target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanAttackNpc(Index, Target, True) Then
                                SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, Target, Vital, spellnum, mapnum
                        End If
                    End If
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.mp, GetPlayerVital(Index, Vitals.mp) - MPCost)
        Call SendVital(Index, Vitals.mp)
        TempPlayer(Index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        Call SendCooldown(Index, spellslot)
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal damage As Long, ByVal spellnum As Long)
On Error Resume Next
Dim sSymbol As String * 1
Dim colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.hp Then colour = BrightGreen
            If Vital = Vitals.mp Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation GetPlayerMap(Index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & damage, colour, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        If increment Then SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + damage
        If Not increment Then SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - damage
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal Index As Long, ByVal damage As Long, ByVal spellnum As Long, ByVal mapnum As Long)
On Error Resume Next
Dim sSymbol As String * 1
Dim colour As Long

    If damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.hp Then colour = BrightGreen
            If Vital = Vitals.mp Then colour = BrightBlue
        Else
            sSymbol = "-"
            colour = Blue
        End If
    
        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg mapnum, sSymbol & damage, colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(Index).x * 32, MapNpc(mapnum).NPC(Index).y * 32
        If increment Then MapNpc(mapnum).NPC(Index).Vital(Vital) = MapNpc(mapnum).NPC(Index).Vital(Vital) + damage
        If Not increment Then MapNpc(mapnum).NPC(Index).Vital(Vital) = MapNpc(mapnum).NPC(Index).Vital(Vital) - damage
    End If
End Sub

Public Sub StunPlayer(ByVal Index As Long, ByVal spellnum As Long)
On Error Resume Next
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(Index).StunDuration = Spell(spellnum).StunDuration
        TempPlayer(Index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned Index
        ' tell him he's stunned
        PlayerMsg Index, "You have been stunned.", BrightRed
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal mapnum As Long, ByVal spellnum As Long)
On Error Resume Next
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(Index).StunDuration = Spell(spellnum).StunDuration
        MapNpc(mapnum).NPC(Index).StunTimer = GetTickCount
    End If
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    Dim ShopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > map(mapnum).MaxX Then x = map(mapnum).MaxX
    If y > map(mapnum).MaxY Then y = map(mapnum).MaxY
    TempPlayer(Index).Target = 0
    TempPlayer(Index).TargetType = TARGET_TYPE_NONE

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If OldMap <> mapnum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    
    If OldMap = mapnum Then
    Else
    SendMusicToOne Index
    End If
    
    
    ' send equipment of all people on new map
    If GetTotalMapPlayers(mapnum) > 0 Then
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) Then
                If GetPlayerMap(i) = mapnum Then
                    SendMapEquipmentTo i, Index
                End If
            End If
        Next
    End If

    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO

        ' Regenerate all NPCs' health
        For i = 1 To MAX_MAP_NPCS

            If MapNpc(OldMap).NPC(i).Num > 0 Then
                MapNpc(OldMap).NPC(i).Vital(Vitals.hp) = GetNpcMaxVital(MapNpc(OldMap).NPC(i).Num, Vitals.hp)
            End If

        Next

    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES
    TempPlayer(Index).GettingMap = YES
    Set buffer = New clsBuffer
    buffer.WriteInteger SCheckForMap
    buffer.WriteLong mapnum
    buffer.WriteLong map(mapnum).Revision
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
    SendFlashlight Index, GetMapFlashlight(mapnum)
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
On Error Resume Next
  Dim buffer As clsBuffer
    Dim mapnum As Long
    Dim gymnum As Long
    Dim x As Long
    Dim y As Long
    Dim Moved As Byte
    Dim MovedSoFar As Boolean
    Dim NewMapX As Byte, NewMapY As Byte
     Movement = MOVING_WALKING
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Moved = NO
    mapnum = GetPlayerMap(Index)
    
    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_BLOCKED Then
                
                    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_RESOURCE Then

                        ' Check to see if the tile is a key and if it is check if its opened
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_KEY Or (map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) - 1) = YES) Then
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type <> TILE_TYPE_GYMBLOCK Then
                            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                        Else
                            
                            gymnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Data1
                            
                            If player(Index).Bedages(gymnum) > GYM_UNDEFEATED Then
                            Call SetPlayerY(Index, GetPlayerY(Index) - 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            Else
                           
                            End If
                        End If
                    End If
                End If
                End If
            Else

                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(Index)).Up > 0 Then
                    NewMapY = map(map(GetPlayerMap(Index)).Up).MaxY
                    Dim nextmapup As Long
                    nextmapup = map(GetPlayerMap(Index)).Up
                    If map(nextmapup).Tile(GetPlayerX(Index), NewMapY).Type = TILE_TYPE_BLOCKED Then
                    PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
                    Else
                    Call PlayerWarp(Index, map(GetPlayerMap(Index)).Up, GetPlayerX(Index), NewMapY)
                    End If
                    
                    
                    Moved = YES
                End If
                

            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If GetPlayerY(Index) < map(mapnum).MaxY Then

                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_BLOCKED Then
                    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_RESOURCE Then

                        ' Check to see if the tile is a key and if it is check if its opened
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_KEY Or (map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index), GetPlayerY(Index) + 1) = YES) Then
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type <> TILE_TYPE_GYMBLOCK Then
                            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            Else
                            
                            gymnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Data1
                            
                            If player(Index).Bedages(gymnum) > GYM_UNDEFEATED Then
                            Call SetPlayerY(Index, GetPlayerY(Index) + 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            Else
                            
                            End If
                        End If
                    End If
                End If
              End If
            Else

                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(Index)).Down > 0 Then
                    Dim nextmapdown As Long
                    nextmapdown = map(GetPlayerMap(Index)).Down
                    If map(nextmapdown).Tile(GetPlayerX(Index), 0).Type = TILE_TYPE_BLOCKED Then
                    PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
                    Else
                    Call PlayerWarp(Index, map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    End If
                    
                    Moved = YES
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) > 0 Then

                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Then

                        ' Check to see if the tile is a key and if it is check if its opened
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) - 1, GetPlayerY(Index)) = YES) Then
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type <> TILE_TYPE_GYMBLOCK Then
                            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            Else
                            gymnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Data1
                            
                            If player(Index).Bedages(gymnum) > GYM_UNDEFEATED Then
                            Call SetPlayerX(Index, GetPlayerX(Index) - 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            End If
                        End If
                    End If
                End If
            End If
            Else

                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(Index)).Left > 0 Then
                    NewMapX = map(map(GetPlayerMap(Index)).Left).MaxX
                    Dim nextmapleft As Long
                    nextmapleft = map(GetPlayerMap(Index)).Left
                    If map(nextmapleft).Tile(NewMapX, GetPlayerY(Index)).Type = TILE_TYPE_BLOCKED Then
                    PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
                    Moved = YES
                    Else
                    Call PlayerWarp(Index, map(GetPlayerMap(Index)).Left, NewMapX, GetPlayerY(Index))
                    Moved = YES
                    End If
                    
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(Index) < map(mapnum).MaxX Then

                ' Check to make sure that the tile is walkable
                If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_BLOCKED Then
                    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_RESOURCE Then

                        ' Check to see if the tile is a key and if it is check if its opened
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_KEY Or (map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(GetPlayerX(Index) + 1, GetPlayerY(Index)) = YES) Then
                        If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type <> TILE_TYPE_GYMBLOCK Then
                            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            Else
                            gymnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Data1
                            
                            If player(Index).Bedages(gymnum) > GYM_UNDEFEATED Then
                            Call SetPlayerX(Index, GetPlayerX(Index) + 1)
                            Set buffer = New clsBuffer

                            With buffer
                                .WriteInteger SPlayerMove
                                .WriteLong Index
                                .WriteLong GetPlayerX(Index)
                                .WriteLong GetPlayerY(Index)
                                .WriteLong GetPlayerDir(Index)
                                .WriteLong Movement
                                SendDataToMapBut Index, GetPlayerMap(Index), .ToArray()
                            End With

                            Set buffer = Nothing
                            Moved = YES
                            End If
                        End If
                    End If
                End If
            End If
            Else

                ' Check to see if we can move them to the another map
                If map(GetPlayerMap(Index)).Right > 0 Then
                    Dim nextmapright As Long
                    nextmapright = map(GetPlayerMap(Index)).Right
                    If map(nextmapright).Tile(0, GetPlayerY(Index)).Type = TILE_TYPE_BLOCKED Then
                    PlayerWarp Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
                    Moved = YES
                    Else
                    Call PlayerWarp(Index, map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Moved = YES
                    End If
                    
                    
                End If
            End If

    End Select

    ' Check to see if the tile is a warp tile, and if so warp them
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        mapnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        x = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
        'TempPlayer(Index).CanPlayerMove = 1
        Call PlayerWarp(Index, mapnum, x, y)
        If TempPlayer(Index).HasBike = YES Then UseBike Index
        Moved = YES
    End If

    ' Check to see if the tile is a door tile, and if so warp them
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
        mapnum = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        x = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        y = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3
        ' send the animation to the map
        SendDoorAnimation GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
        'TempPlayer(Index).CanPlayerMove = 1
        Call PlayerWarp(Index, mapnum, x, y)
        Moved = YES
    End If

    ' Check for key trigger open
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        x = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        y = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

        If map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
            Set buffer = New clsBuffer
            buffer.WriteInteger SMapKey
            buffer.WriteLong x
            buffer.WriteLong y
            buffer.WriteByte 1
            SendDataToMap GetPlayerMap(Index), buffer.ToArray()
            Set buffer = Nothing
            Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
        End If
    End If
    
    ' Check for a shop, and if so open it
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SHOP Then
        x = map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        If x > 0 Then ' shop exists?
            If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                Set buffer = New clsBuffer
                buffer.WriteInteger SOpenShop ' send packet opening the shop
                buffer.WriteLong x
                SendDataTo Index, buffer.ToArray()
                Set buffer = Nothing
                TempPlayer(Index).InShop = x ' stops movement and the like
            End If
        End If
    End If
    
    ' check for battle and enter them into a battle if it's true
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_BATTLE Then
    If TempPlayer(Index).PokemonBattle.PokemonNumber = 0 Then
    If CheckPlayerDefeat(Index) = False Then
    SendAnimation player(Index).map, 1, player(Index).x, player(Index).y - 0.2
    Dim na As Long
    na = Rand(1, 100)
    If na <= 25 Then initBattle Index 'there is 20% of chance for pokemon to spawn
    End If
    End If
    End If
    
    ' check for pokemon healing tile
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
      If HealPokemons(Index) Then
      SetPlayerSpawn Index, GetPlayerX(Index), GetPlayerY(Index), GetPlayerMap(Index)
      End If
    
    End If
    'check for Custom Script
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CUSTOMSCRIPT Then
    Call CustomScript(Index, map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
    End If
    'check for pokemon spawn tile
     If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPAWN Then
       SetPlayerSpawn Index, GetPlayerX(Index), GetPlayerY(Index), GetPlayerMap(Index)
    End If
    'check for pokemon storage tile
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_STORAGE Then
       'We will update storage
       Call SendPlayerStorage(Index)
       'And the open it
       Call SendOpenStorage(Index)
       
       Call SendPlayerStorage(Index)
    End If
    
    'Check for Bank Tile
    If map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_BANK Then
       SendPlayerData (Index)
       SendUpdateBank (Index)
       SendOpenBank (Index)
    End If
    
    
    
    'Steps
    If Moved = YES Then
    TempPlayer(Index).eggStepsTemp = TempPlayer(Index).eggStepsTemp + 1
    
    End If
    
    ' They tried to hack
    If Moved = NO Then
        Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    End If

End Sub

Sub SetPlayerSpawn(ByVal Index As Long, ByVal x As Long, ByVal y As Long, map As Long)
On Error Resume Next
player(Index).SX = x
player(Index).SY = y
player(Index).SMap = map
PlayerMsg Index, "Spawn saved!", Yellow
End Sub

Sub SetPlayerMood(ByVal Index As Long, ByVal moodState As Long)
On Error Resume Next
If Index > MAX_PLAYERS Then Exit Sub
player(Index).mood = moodState
End Sub

Sub SpawnPlayer(ByVal Index As Long)
On Error Resume Next
PlayerWarp Index, player(Index).SMap, player(Index).SX, player(Index).SY
PlayerMsg Index, "You have been spawned in " & map(player(Index).map).Name, Yellow
End Sub
Sub CheckEquippedItems(ByVal Index As Long)
On Error Resume Next
    Dim slot As Long
    Dim itemNum As Long
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        itemNum = GetPlayerEquipment(Index, i)

        If itemNum > 0 Then

            Select Case i
                Case Equipment.Weapon

                    If item(itemNum).Type <> ITEM_TYPE_WEAPON Then SetPlayerEquipment Index, 0, i
                Case Equipment.Armor

                    If item(itemNum).Type <> ITEM_TYPE_ARMOR Then SetPlayerEquipment Index, 0, i
                Case Equipment.Helmet

                    If item(itemNum).Type <> ITEM_TYPE_HELMET Then SetPlayerEquipment Index, 0, i
                Case Equipment.Shield

                    If item(itemNum).Type <> ITEM_TYPE_SHIELD Then SetPlayerEquipment Index, 0, i
               Case Equipment.Mask

                    If item(itemNum).Type <> ITEM_TYPE_MASK Then SetPlayerEquipment Index, 0, i
                    Case Equipment.Outfit

                    If item(itemNum).Type <> ITEM_TYPE_OUTFIT Then SetPlayerEquipment Index, 0, i
                    
            End Select

        Else
            SetPlayerEquipment Index, 0, i
        End If

    Next

End Sub

Function FindOpenInvSlot(ByVal Index As Long, ByVal itemNum As Long) As Long
On Error Resume Next
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    'If item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV

            If GetPlayerInvItemNum(Index, i) = itemNum Then
                FindOpenInvSlot = i
                Exit Function
           End If
        Next

    'End If

    For i = 1 To MAX_INV

         'Try to find an open free slot
        If GetPlayerInvItemNum(Index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If

    Next

End Function

Function HasItem(ByVal Index As Long, ByVal itemNum As Long) As Long
On Error Resume Next
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Function
    End If

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemNum Then
            If item(itemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(Index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If

    Next

End Function


Function GetItemSlot(ByVal Index As Long, ByVal item As Long)
On Error Resume Next
Dim i As Long
For i = 1 To MAX_INV
If GetPlayerInvItemNum(Index, i) = item Then
GetItemSlot = i
Exit Function
End If
Next


End Function

Sub TakeItem(ByVal Index As Long, ByVal itemNum As Long, ByVal ItemVal As Long)
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Dim TakeItem As Boolean

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If ItemVal < 1 Then ItemVal = 1

    For i = 1 To MAX_INV

        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, i) = itemNum Then
            If item(itemNum).Type = ITEM_TYPE_CURRENCY Then

                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                    SendInventory Index
                    Call SendInventoryUpdate(Index, i)
                    PlayerMsg Index, "You lost " & ItemVal & " " & Trim$(item(itemNum).Name), BrightRed
                End If

            Else
                TakeItem = True
            End If

            If TakeItem = True Then
            If GetPlayerInvItemValue(Index, i) <= ItemVal Then
                Call SetPlayerInvItemNum(Index, i, 0)
                Call SetPlayerInvItemValue(Index, i, 0)
                ' Send the inventory update
                Call SendInventoryUpdate(Index, i)
                PlayerMsg Index, "You lost all " & Trim$(item(itemNum).Name) & "'s", BrightRed
                Exit Sub
            Else
                Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) - ItemVal)
                SendInventory Index
                Call SendInventoryUpdate(Index, i)
                PlayerMsg Index, "You lost " & ItemVal & " " & Trim$(item(itemNum).Name), BrightRed
                Exit Sub
            End If
            End If
        End If

    Next
SendInventory Index
End Sub

Function GiveItem(ByVal Index As Long, ByVal itemNum As Long, ByVal ItemVal As Long, Optional msg As Long = YES) As Boolean
On Error Resume Next
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or itemNum <= 0 Or itemNum > MAX_ITEMS Then
        GiveItem = False
        Exit Function
    End If

    i = FindOpenInvSlot(Index, itemNum)
    If ItemVal < 1 Then ItemVal = 1
    
   
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(Index, i, itemNum)
        Call SetPlayerInvItemValue(Index, i, GetPlayerInvItemValue(Index, i) + ItemVal)
        Call SendInventoryUpdate(Index, i)
        GiveItem = True
        If msg = YES Then
        PlayerMsg Index, "You earned " & ItemVal & " " & Trim$(item(itemNum).Name) & " ( " & GetPlayerInvItemValue(Index, i) & " in inventory)!", BrightGreen
        End If
    Else
        Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
        GiveItem = False
    End If
    SendInventory Index
    
    
End Function

Function HasSpell(ByVal Index As Long, ByVal spellnum As Long) As Boolean
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = spellnum Then
            HasSpell = True
            Exit Function
        End If

    Next

End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS

        If GetPlayerSpell(Index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If

    Next

End Function

Sub PlayerMapGetItem(ByVal Index As Long)
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Dim mapnum As Long
    Dim msg As String

    If Not IsPlaying(Index) Then Exit Sub
    mapnum = GetPlayerMap(Index)

    For i = 1 To MAX_MAP_ITEMS

        ' See if theres even an item here
        If (MapItem(mapnum, i).Num > 0) Then
            If (MapItem(mapnum, i).Num <= MAX_ITEMS) Then

                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(Index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(Index)) Then
                        ' Find open slot
                        n = FindOpenInvSlot(Index, MapItem(mapnum, i).Num)

                        ' Open slot available?
                        If n <> 0 Then
                            ' Set item in players inventor
                            Call SetPlayerInvItemNum(Index, n, MapItem(mapnum, i).Num)

                            If item(GetPlayerInvItemNum(Index, n)).Type = ITEM_TYPE_CURRENCY Then
                                Call SetPlayerInvItemValue(Index, n, GetPlayerInvItemValue(Index, n) + MapItem(mapnum, i).value)
                                msg = MapItem(mapnum, i).value & " " & Trim$(item(GetPlayerInvItemNum(Index, n)).Name)
                            Else
                                Call SetPlayerInvItemValue(Index, n, 0)
                                msg = CheckGrammar(Trim$(item(GetPlayerInvItemNum(Index, n)).Name), 1)
                            End If

                            ' Erase item from the map
                            MapItem(mapnum, i).Num = 0
                            MapItem(mapnum, i).value = 0
                            MapItem(mapnum, i).x = 0
                            MapItem(mapnum, i).y = 0
                            Call SendInventoryUpdate(Index, n)
                            Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), 0, 0)
                            'Call PlayerMsg(Index, Msg, Yellow)
                            SendActionMsg GetPlayerMap(Index), msg, White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                            Exit For
                        Else
                            Call PlayerMsg(Index, "Your inventory is full.", BrightRed)
                            Exit For
                        End If
                    End If
                End If
            End If
        End If

    Next

End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal amount As Long)
On Error Resume Next
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
                If item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                Else
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                End If
                Call SendInventoryUpdate(Index, InvNum)
            End If
        End If


End Sub

Sub CheckPlayerLevelUp(ByVal Index As Long)

On Error Resume Next
Exit Sub
    Dim i As Long
    Dim expRollover As Long
    Dim level_count As Long
    
    level_count = 0
    
    Do While GetPlayerExp(Index) >= GetPlayerNextLevel(Index)
        expRollover = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
        Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + 3)
        Call SetPlayerExp(Index, expRollover)
        level_count = level_count + 1
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            'singular
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " level!", Brown
        Else
            'plural
            GlobalMsg GetPlayerName(Index) & " has gained " & level_count & " levels!", Brown
        End If
        SendEXP Index
        SendPlayerData Index
    End If
End Sub

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case hp
            i = (GetPlayerStat(Index, Stats.vitality) \ 2)
        Case mp
            i = (GetPlayerStat(Index, Stats.spirit) \ 2)
        Case SP
            i = (GetPlayerStat(Index, Stats.spirit) \ 2)
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal Index As Long) As String
On Error Resume Next
    GetPlayerLogin = Trim$(player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
On Error Resume Next
    player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
On Error Resume Next
    GetPlayerPassword = Trim$(player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
On Error Resume Next
    player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
On Error Resume Next
    player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
On Error Resume Next
    GetPlayerClass = player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
On Error Resume Next
    player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
On Error Resume Next
    player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = player(Index).level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal level As Long)
On Error Resume Next
    If level > MAX_LEVELS Then Exit Sub
    player(Index).level = level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
On Error Resume Next
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerStat(Index, Stats.strength) + GetPlayerStat(Index, Stats.endurance) + GetPlayerStat(Index, Stats.intelligence) + GetPlayerStat(Index, Stats.spirit) + GetPlayerPOINTS(Index)) * 25
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
On Error Resume Next
    GetPlayerExp = player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
On Error Resume Next
    player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
On Error Resume Next
    player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
On Error Resume Next
    player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal value As Long)
On Error Resume Next
    player(Index).Vital(Vital) = value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        player(Index).Vital(Vital) = 0
    End If

End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function

    Select Case Vital
        Case hp
            GetPlayerMaxVital = (player(Index).level + (GetPlayerStat(Index, Stats.vitality) \ 2) + Class(player(Index).Class).Stat(Stats.vitality)) * 2
        Case mp
            GetPlayerMaxVital = (player(Index).level + (GetPlayerStat(Index, Stats.intelligence) \ 2) + Class(player(Index).Class).Stat(Stats.intelligence)) * 2
        Case SP
            GetPlayerMaxVital = (player(Index).level + (GetPlayerStat(Index, Stats.spirit) \ 2) + Class(player(Index).Class).Stat(Stats.spirit)) * 2
    End Select

End Function

Public Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
On Error Resume Next
    Dim x As Long, i As Long
    If Index > MAX_PLAYERS Then Exit Function
    
    x = player(Index).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If player(Index).Equipment(i) > 0 Then
            If item(player(Index).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + item(player(Index).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Public Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerRawStat = player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal value As Long)
On Error Resume Next
    player(Index).Stat(Stat) = value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
On Error Resume Next
    player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = player(Index).map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    If mapnum > 0 And mapnum <= MAX_MAPS Then
        player(Index).map = mapnum
    End If

End Sub

Function GetPlayerX(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
On Error Resume Next
    player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
On Error Resume Next
    player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
On Error Resume Next
    player(Index).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemNum = player(Index).Inv(invslot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal itemNum As Long)
On Error Resume Next
    player(Index).Inv(invslot).Num = itemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = player(Index).Inv(invslot).value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal ItemValue As Long)
On Error Resume Next
    player(Index).Inv(invslot).value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long) As Long
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = player(Index).Spell(spellslot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal spellslot As Long, ByVal spellnum As Long)
On Error Resume Next
    player(Index).Spell(spellslot) = spellnum
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
On Error Resume Next
    player(Index).Equipment(EquipmentSlot) = InvNum
End Sub

' ToDo
Sub OnDeath(ByVal Index As Long)
On Error Resume Next
    Dim i As Long
    
    ' Set HP to nothing
    Call SetPlayerVital(Index, Vitals.hp, 0)

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(Index, i) > 0 Then
            PlayerMapDropItem Index, GetPlayerEquipment(Index, i), 0
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(Index, DIR_DOWN)
    Call PlayerWarp(Index, START_MAP, START_X, START_Y)
    
    ' Clear spell casting
    TempPlayer(Index).SpellBuffer = 0
    TempPlayer(Index).SpellBufferTimer = 0
    'Call SendClearSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(Index, Vitals.hp, GetPlayerMaxVital(Index, Vitals.hp))
    Call SetPlayerVital(Index, Vitals.mp, GetPlayerMaxVital(Index, Vitals.mp))
    Call SetPlayerVital(Index, Vitals.SP, GetPlayerMaxVital(Index, Vitals.SP))
    Call SendVital(Index, Vitals.hp)
    Call SendVital(Index, Vitals.mp)
    Call SendVital(Index, Vitals.SP)

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(Index) = YES Then
        Call SetPlayerPK(Index, NO)
        Call SendPlayerData(Index)
    End If

End Sub

Sub CheckResource(ByVal Index As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim damage As Long
    
    If map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        Resource_num = 0
        Resource_index = map(GetPlayerMap(Index)).Tile(x, y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count

            If ResourceCache(GetPlayerMap(Index)).ResourceData(i).x = x Then
                If ResourceCache(GetPlayerMap(Index)).ResourceData(i).y = y Then
                    Resource_num = i
                End If
            End If

        Next

        If Resource_num > 0 Then
            If GetPlayerEquipment(Index, Weapon) > 0 Then
                If item(GetPlayerEquipment(Index, Weapon)).Data3 = Resource(Resource_index).ToolRequired Then

                    ' inv space?
                    If Resource(Resource_index).ItemReward > 0 Then
                        If FindOpenInvSlot(Index, Resource(Resource_index).ItemReward) = 0 Then
                            PlayerMsg Index, "You have no inventory space.", BrightRed
                            Exit Sub
                        End If
                    End If

                    ' check if already cut down
                    If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 0 Then
                    
                        rX = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).x
                        rY = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).y
                        
                        damage = item(GetPlayerEquipment(Index, Weapon)).Data2
                    
                        ' check if damage is more than health
                        If damage > 0 Then
                            ' cut it down!
                            If ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - damage <= 0 Then
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
                                SendResourceCacheToMap GetPlayerMap(Index), Resource_num
                                SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                                GiveItem Index, Resource(Resource_index).ItemReward, 1
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            Else
                                ' just do the damage
                                ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(Index)).ResourceData(Resource_num).cur_health - damage
                                SendActionMsg GetPlayerMap(Index), "-" & damage, BrightRed, 1, (rX * 32), (rY * 32)
                                SendAnimation GetPlayerMap(Index), Resource(Resource_index).Animation, rX, rY
                            End If
                        Else
                            ' too weak
                            SendActionMsg GetPlayerMap(Index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
                        End If
                    Else
                        SendActionMsg GetPlayerMap(Index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                    End If

                Else
                    PlayerMsg Index, "You have the wrong type of tool equiped.", BrightRed
                End If

            Else
                PlayerMsg Index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Public Sub GivePokemon(ByVal Index As Long, ByVal pokemonnum As Long, Optional level As Long = 1, Optional ByVal shiny As Long = 0, Optional ByVal tradeable As Long = YES, Optional ByVal customNature As Long = 0, Optional ByVal extraTP As Long = 0)
On Error Resume Next
Dim x As Long
Dim i As Long
Dim y As Long
Dim nat As Long
nat = Rand(1, MAX_NATURES)
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    For x = 1 To 6
        If player(Index).PokemonInstance(x).PokemonNumber = 0 Then
        
            ' give actual pokemon num
            player(Index).PokemonInstance(x).PokemonNumber = pokemonnum
            
            'heal it up
            player(Index).PokemonInstance(x).hp = CalculateStat(Pokemon(pokemonnum).MaxHp, STAT_HP)
            player(Index).PokemonInstance(x).MaxHp = CalculateStat(Pokemon(pokemonnum).MaxHp, STAT_HP)
            player(Index).PokemonInstance(x).atk = CalculateStat(Pokemon(pokemonnum).atk, STAT_ATK)
            player(Index).PokemonInstance(x).def = CalculateStat(Pokemon(pokemonnum).def, STAT_DEF)
            player(Index).PokemonInstance(x).spatk = CalculateStat(Pokemon(pokemonnum).spatk, STAT_SPATK)
            player(Index).PokemonInstance(x).spdef = CalculateStat(Pokemon(pokemonnum).spdef, STAT_SPDEF)
            player(Index).PokemonInstance(x).spd = CalculateStat(Pokemon(pokemonnum).spd, STAT_SPEED)
            ' set values from base values
           
            player(Index).PokemonInstance(x).Happiness = Pokemon(pokemonnum).Happiness
            player(Index).PokemonInstance(x).isShiny = shiny
            player(Index).PokemonInstance(x).isTradeable = tradeable
            'Check for random tps
            If level > 1 Then
Dim availableTP As Long
availableTP = level * 3 - 3
Do While availableTP > 0
Dim stattoadd As Long
stattoadd = Rand(1, 6)
Select Case stattoadd
Case STAT_ATK
player(Index).PokemonInstance(x).atk = player(Index).PokemonInstance(x).atk + 1
Case STAT_DEF
player(Index).PokemonInstance(x).def = player(Index).PokemonInstance(x).def + 1
Case STAT_SPATK
player(Index).PokemonInstance(x).spatk = player(Index).PokemonInstance(x).spatk + 1
Case STAT_SPDEF
player(Index).PokemonInstance(x).spdef = player(Index).PokemonInstance(x).spdef + 1
Case STAT_SPEED
player(Index).PokemonInstance(x).spd = player(Index).PokemonInstance(x).spd + 1
Case STAT_HP
player(Index).PokemonInstance(x).MaxHp = player(Index).PokemonInstance(x).MaxHp + 2
End Select
availableTP = availableTP - 1
Loop
            End If
            
            ' reset some values
            player(Index).PokemonInstance(x).hp = player(Index).PokemonInstance(x).MaxHp
            player(Index).PokemonInstance(x).level = level
            player(Index).PokemonInstance(x).EXP = 0
            If shiny = YES Then
             player(Index).PokemonInstance(x).TP = 20 + extraTP
            Else
             player(Index).PokemonInstance(x).TP = 0 + extraTP
            End If
           
            Dim a As Long
            Dim b As Long
            For a = 1 To 4
            If GetPokemonMove(pokemonnum, a) > 0 Then
            player(Index).PokemonInstance(x).moves(a).number = GetPokemonMove(pokemonnum, a)
            player(Index).PokemonInstance(x).moves(a).pp = PokemonMove(GetPokemonMove(pokemonnum, a)).pp
            Else
            player(Index).PokemonInstance(x).moves(a).number = 0
            player(Index).PokemonInstance(x).moves(a).pp = 0
            End If
            Next
            

            'If player(index).PokemonInstance(x).moves(1).number = 0 Then
            'player(index).PokemonInstance(x).moves(1).pp = 0
            'Else
            'player(index).PokemonInstance(x).moves(1).pp = PokemonMove(player(index).PokemonInstance(x).moves(1).number).pp
            'End If
            If customNature > 0 Then
             player(Index).PokemonInstance(x).nature = customNature
            Else
            player(Index).PokemonInstance(x).nature = nat
            End If
            
            'Removed because of registration
            
            'send message
            PlayerMsg Index, "You have received " & CheckGrammar(Trim$(Pokemon(pokemonnum).Name)) & "!", BrightCyan
            ' update player
            SendPlayerPokemon Index
            Exit Sub
        End If
    Next
    
    For y = 1 To 250
    If player(Index).StoragePokemonInstance(y).PokemonNumber <= 1 Then
    ' give actual pokemon num
            player(Index).StoragePokemonInstance(y).PokemonNumber = pokemonnum
            
            'heal it up
            player(Index).StoragePokemonInstance(y).hp = CalculateStat(Pokemon(pokemonnum).MaxHp, STAT_HP)
            player(Index).StoragePokemonInstance(y).MaxHp = CalculateStat(Pokemon(pokemonnum).MaxHp, STAT_HP)
            player(Index).StoragePokemonInstance(y).atk = CalculateStat(Pokemon(pokemonnum).atk, STAT_ATK)
            player(Index).StoragePokemonInstance(y).def = CalculateStat(Pokemon(pokemonnum).def, STAT_DEF)
            player(Index).StoragePokemonInstance(y).spatk = CalculateStat(Pokemon(pokemonnum).spatk, STAT_SPATK)
            player(Index).StoragePokemonInstance(y).spdef = CalculateStat(Pokemon(pokemonnum).spdef, STAT_SPDEF)
            player(Index).StoragePokemonInstance(y).spd = CalculateStat(Pokemon(pokemonnum).MaxHp, STAT_SPEED)
            
            player(Index).StoragePokemonInstance(y).pp = Pokemon(pokemonnum).maxpp
            ' set values from base values
        
                       'Check for random tps
            If level > 1 Then
Dim availableTP2 As Long
availableTP2 = level * 3 - 3
Do While availableTP2 > 0
Dim stattoadd2 As Long
stattoadd2 = Rand(1, 6)
Select Case stattoadd2
Case STAT_ATK
player(Index).StoragePokemonInstance(y).atk = player(Index).StoragePokemonInstance(y).atk + 1
Case STAT_DEF
player(Index).StoragePokemonInstance(y).def = player(Index).StoragePokemonInstance(y).def + 1
Case STAT_SPATK
player(Index).StoragePokemonInstance(y).spatk = player(Index).StoragePokemonInstance(y).spatk + 1
Case STAT_SPDEF
player(Index).StoragePokemonInstance(y).spdef = player(Index).StoragePokemonInstance(y).spdef + 1
Case STAT_SPEED
player(Index).StoragePokemonInstance(y).spd = player(Index).StoragePokemonInstance(y).spd + 1
Case STAT_HP
player(Index).StoragePokemonInstance(y).MaxHp = player(Index).StoragePokemonInstance(y).MaxHp + 2
End Select
availableTP2 = availableTP2 - 1
Loop
            End If
            
            player(Index).StoragePokemonInstance(y).hp = player(Index).StoragePokemonInstance(y).MaxHp
            player(Index).StoragePokemonInstance(y).Happiness = Pokemon(pokemonnum).Happiness
            player(Index).StoragePokemonInstance(y).isShiny = shiny
            player(Index).StoragePokemonInstance(y).isTradeable = tradeable
            ' reset some values
            player(Index).StoragePokemonInstance(y).level = level
            player(Index).StoragePokemonInstance(y).EXP = 0
            If shiny = YES Then
            player(Index).StoragePokemonInstance(y).TP = 20 + extraTP
            Else
            player(Index).StoragePokemonInstance(y).TP = 0 + extraTP
            End If
            

            For a = 1 To 4
            If GetPokemonMove(pokemonnum, a) > 0 Then
            player(Index).StoragePokemonInstance(y).moves(a).number = GetPokemonMove(pokemonnum, a)
            player(Index).StoragePokemonInstance(y).moves(a).pp = PokemonMove(GetPokemonMove(pokemonnum, a)).pp
            Else
            player(Index).StoragePokemonInstance(y).moves(a).number = 0
            player(Index).StoragePokemonInstance(y).moves(a).pp = 0
            End If
            Next
            If customNature > 0 Then
            player(Index).StoragePokemonInstance(y).nature = customNature
            Else
            player(Index).StoragePokemonInstance(y).nature = nat
            End If
            PlayerMsg Index, Trim$(Pokemon(pokemonnum).Name) & " is stored in storage! - Slot " & y, Yellow
            Exit Sub
    End If
    Next
End Sub

Public Sub TakePokemon(ByVal Index As Long, ByVal pokemonnum As Long)
On Error Resume Next
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If pokemonnum < 1 Or pokemonnum > 6 Then Exit Sub
    
    If player(Index).PokemonInstance(pokemonnum).PokemonNumber > 0 Then
        ' send message (early to give name)
        'PlayerMsg index, "You had " & Trim$(Pokemon(Player(index).PokemonInstance(pokemonnum).PokemonNumber).Name) & " taken away by an admin.", BrightRed
        ' take away
        player(Index).PokemonInstance(pokemonnum).PokemonNumber = 0
        'update player
        SendPlayerPokemon Index
        Exit Sub
    End If
End Sub

Public Sub TakePlayerPokemon(ByVal Index As Long, ByVal slot As Long)
On Error Resume Next
Dim newPoke As PokemonInstanceRec
player(Index).PokemonInstance(slot) = newPoke
End Sub

Public Sub DepositPokemon(ByVal Index As Long, ByVal pokemonnum As Long)
On Error Resume Next
Dim emptyslot As Long
Dim i As Long
Dim a As Long
Dim b As Long
Dim emptypokemons As Long
'Find empty slot
For i = 1 To 250
If player(Index).StoragePokemonInstance(i).PokemonNumber <= 0 Then
emptyslot = i
Exit For
End If
Next
'Check if its only pokemon
For a = 1 To 6
If player(Index).PokemonInstance(a).PokemonNumber = 0 Then
emptypokemons = emptypokemons + 1
End If
Next

'Deposit
If emptypokemons < 5 Then
Else
PlayerMsg Index, "This is your only pokemon!", BrightRed
Exit Sub
End If

If emptyslot > 0 Then
If pokemonnum > 0 And pokemonnum <= 6 Then
player(Index).StoragePokemonInstance(emptyslot) = player(Index).PokemonInstance(pokemonnum)
If pokemonnum = 1 Then
For b = 2 To 6
If player(Index).PokemonInstance(b).PokemonNumber > 0 Then
player(Index).PokemonInstance(1) = player(Index).PokemonInstance(b)
Call SendPlayerPokemon(Index)
Call SendPlayerStorage(Index)
TakePokemon Index, b
Exit For
End If
Next
Else
Call SendPlayerPokemon(Index)
Call SendPlayerStorage(Index)
TakePokemon Index, pokemonnum
End If
End If
Else
PlayerMsg Index, "Your storage is full!", BrightRed
End If

End Sub

Sub WithdrawPokemon(Index As Long, ByVal storageslot As Long)
On Error Resume Next
Dim emptyslot As Long
Dim i As Long

For i = 1 To 6
If player(Index).PokemonInstance(i).PokemonNumber = 0 Then
emptyslot = i
Exit For
End If
Next


If emptyslot > 0 Then
player(Index).PokemonInstance(emptyslot) = player(Index).StoragePokemonInstance(storageslot)
player(Index).StoragePokemonInstance(storageslot).PokemonNumber = 0
SendPlayerPokemon Index
SendPlayerStorage Index
SendStoragePokemonLoad Index, storageslot
Else
PlayerMsg Index, "Your main pokemon slots are full!", BrightRed
End If
End Sub


Public Sub initBattle(ByVal Index As Long)
On Error Resume Next
ResetBattlePokemon (Index)
Dim i As Long
Dim x As Long
Dim wildpoke As Long
Dim slot As Long
Dim frmlvl As Long
Dim tolvl As Long
Dim cstm As Long
Dim slt As Long

If TempPlayer(Index).PokemonBattle.PokemonNumber > 0 Then Exit Sub

If player(Index).PokemonInstance(1).hp > 0 And player(Index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(Index).PokemonInstance(i).PokemonNumber > 0 Then
If player(Index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(Index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If

If slot < 1 Then Exit Sub


'Set wild pokemon
For i = 1 To MAX_MAP_POKEMONS
If SpawnChance(map(GetPlayerMap(Index)).Pokemon(i).Chance) = True Then
wildpoke = map(GetPlayerMap(Index)).Pokemon(i).PokemonNumber
frmlvl = map(GetPlayerMap(Index)).Pokemon(i).LevelFrom
tolvl = map(GetPlayerMap(Index)).Pokemon(i).LevelTo
cstm = map(GetPlayerMap(Index)).Pokemon(i).Custom
slt = i
Exit For
End If
Next


If wildpoke < 1 Or wildpoke > 721 Then Exit Sub 'No battle if there is not pokemon to spawn

'If there is pokemon then we are going to set BattlePokemon ready!

ResetBattlePokemon (Index)
TempPlayer(Index).PokemonBattle.PokemonNumber = wildpoke
TempPlayer(Index).PokemonBattle.level = Rand(frmlvl, tolvl)
TempPlayer(Index).PokemonBattle.MapSlot = slt
TempPlayer(Index).PokemonBattle.nature = Rand(1, MAX_NATURES)
TempPlayer(Index).PokemonBattle.status = STATUS_NOTHING
TempPlayer(Index).PokemonBattle.turnsneed = 0
TempPlayer(Index).PokemonBattle.statusturn = 0
Dim rndNum As Long
rndNum = Rand(1, 100) '100k normal
If GetPlayerAccess(Index) >= ADMIN_DEVELOPER Then
'PlayerMsg index, "Shiny num: " & rndnum, Yellow
End If
If rndNum = 23 Then
TempPlayer(Index).PokemonBattle.isShiny = YES ' for now
GlobalMsg "(SHINY!) " & Trim$(player(Index).Name) & " encountered lvl." & TempPlayer(Index).PokemonBattle.level & " shiny " & Trim$(Pokemon(wildpoke).Name), Pink
End If

'MoveThing for now
TempPlayer(Index).PokemonBattle.moves(1).number = Pokemon(wildpoke).moves(1)
Select Case cstm
Case YES
TempPlayer(Index).PokemonBattle.atk = map(GetPlayerMap(Index)).Pokemon(slt).atk
TempPlayer(Index).PokemonBattle.def = map(GetPlayerMap(Index)).Pokemon(slt).def
TempPlayer(Index).PokemonBattle.spatk = map(GetPlayerMap(Index)).Pokemon(slt).spatk
TempPlayer(Index).PokemonBattle.spdef = map(GetPlayerMap(Index)).Pokemon(slt).spdef
TempPlayer(Index).PokemonBattle.spd = map(GetPlayerMap(Index)).Pokemon(slt).spd
TempPlayer(Index).PokemonBattle.hp = map(GetPlayerMap(Index)).Pokemon(slt).hp
TempPlayer(Index).PokemonBattle.MaxHp = map(GetPlayerMap(Index)).Pokemon(slt).hp
Case NO
TempPlayer(Index).PokemonBattle.atk = CalculateStat(Pokemon(wildpoke).atk, STAT_ATK)
TempPlayer(Index).PokemonBattle.def = CalculateStat(Pokemon(wildpoke).def, STAT_DEF)
TempPlayer(Index).PokemonBattle.spatk = CalculateStat(Pokemon(wildpoke).spatk, STAT_SPATK)
TempPlayer(Index).PokemonBattle.spdef = CalculateStat(Pokemon(wildpoke).spdef, STAT_SPDEF)
TempPlayer(Index).PokemonBattle.spd = CalculateStat(Pokemon(wildpoke).spd, STAT_SPEED)
TempPlayer(Index).PokemonBattle.MaxHp = CalculateStat(Pokemon(wildpoke).MaxHp, STAT_HP)
If TempPlayer(Index).PokemonBattle.level > 1 Then
Dim availableTP As Long
availableTP = TempPlayer(Index).PokemonBattle.level * 3 - 3
Do While availableTP > 0
Dim stattoadd As Long
stattoadd = Rand(1, 6)
Select Case stattoadd
Case STAT_ATK
TempPlayer(Index).PokemonBattle.atk = TempPlayer(Index).PokemonBattle.atk + 1
Case STAT_DEF
TempPlayer(Index).PokemonBattle.def = TempPlayer(Index).PokemonBattle.def + 1
Case STAT_SPATK
TempPlayer(Index).PokemonBattle.spatk = TempPlayer(Index).PokemonBattle.spatk + 1
Case STAT_SPDEF
TempPlayer(Index).PokemonBattle.spdef = TempPlayer(Index).PokemonBattle.spdef + 1
Case STAT_SPEED
TempPlayer(Index).PokemonBattle.spd = TempPlayer(Index).PokemonBattle.spd + 1
Case STAT_HP
TempPlayer(Index).PokemonBattle.MaxHp = TempPlayer(Index).PokemonBattle.MaxHp + 2
End Select
availableTP = availableTP - 1
Loop
End If
TempPlayer(Index).PokemonBattle.hp = TempPlayer(Index).PokemonBattle.MaxHp

End Select

For x = 1 To 6
player(Index).PokemonInstance(x).batk = player(Index).PokemonInstance(x).atk
player(Index).PokemonInstance(x).bdef = player(Index).PokemonInstance(x).def
player(Index).PokemonInstance(x).bspd = player(Index).PokemonInstance(x).spd
player(Index).PokemonInstance(x).bspatk = player(Index).PokemonInstance(x).spatk
player(Index).PokemonInstance(x).bspdef = player(Index).PokemonInstance(x).spdef
Next



'Set turn (My Speed>Enemy Speed = MyTurn)
If player(Index).PokemonInstance(slot).spd > TempPlayer(Index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(Index).BattleTurn = True
Else
If player(Index).PokemonInstance(slot).spd = TempPlayer(Index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(Index).BattleTurn = True
Else
TempPlayer(Index).BattleTurn = False
End If
Else
TempPlayer(Index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
TempPlayer(Index).BattleCurrentTurn = 1
SendNpcBattle Index, slot
Call SendActionMsg(GetPlayerMap(Index), "Encounter: " & Trim$(Pokemon(wildpoke).Name), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)

End Sub

Public Sub RemoveStoragePokemon(ByVal Index As Long, ByVal storagenum As Long)
On Error Resume Next
player(Index).StoragePokemonInstance(storagenum).PokemonNumber = 0
Call SendPlayerStorage(Index)
Call SendPlayerData(Index)
End Sub


Sub SetAsLeader(ByVal Index As Long, ByVal pokeSlot As Long)
On Error Resume Next
Dim currentleader As PokemonInstanceRec
Dim newleader As PokemonInstanceRec
'Check if pokemon is already leader
If pokeSlot = 1 Then Exit Sub
'If not set it as leader
currentleader = player(Index).PokemonInstance(1)
newleader = player(Index).PokemonInstance(pokeSlot)
'CHANGE THE LEADER
player(Index).PokemonInstance(1) = newleader
player(Index).PokemonInstance(pokeSlot) = currentleader
'Send all info
SendPlayerData Index
SendPlayerPokemon Index

'BOOM DONE!
End Sub


Sub SetTrade(ByVal Index As Long, ByVal Name As String)
On Error Resume Next
TempPlayer(Index).TradeName = Name
End Sub


Sub DoTrade(ByVal Index As Long)
On Error Resume Next
Dim trader As Long
Dim fs As Long
Dim i As Long
Dim mp As Long
Dim TP As Long
Dim allowedLevelIndex As Long
Dim allowedLevelTrader As Long
trader = FindPlayer(Trim$(TempPlayer(Index).TradeName))
'------------------------------------------------------------
'Trader check!
If TempPlayer(Index).TradePoke > 0 Then
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).PokemonNumber > 0 Then
fs = 0
For i = 1 To 6
If player(trader).PokemonInstance(i).PokemonNumber = 0 Then
fs = i
Exit For
End If
Next
If fs = 0 Then
PlayerMsg Index, "Your trading partner has 6 pokemons with him!", Yellow
PlayerMsg trader, "You have 6 pokemons with you!", Yellow
Exit Sub
End If
End If
End If

If TempPlayer(Index).TradePoke > 0 Then
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).level > 0 Then
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).isTradeable <> YES Then
PlayerMsg trader, "Your trading partner pokemon is untradeable!", Yellow
PlayerMsg Index, "Your pokemon is untradeable!", Yellow
Exit Sub
End If
End If
End If

If TempPlayer(trader).TradePoke > 0 Then
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).level > 0 Then
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).isTradeable <> YES Then
PlayerMsg Index, "Your trading partner pokemon is untradeable!", Yellow
PlayerMsg trader, "Your pokemon is untradeable!", Yellow
Exit Sub
End If
End If
End If

If TempPlayer(Index).TradeItem > 0 Then
If IsItemTradeable(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem)) = False Then
PlayerMsg trader, "Your trading partner item is untradeable!", Yellow
PlayerMsg Index, "Your item is untradeable!", Yellow
Exit Sub
End If
End If

If TempPlayer(trader).TradeItem > 0 Then
If IsItemTradeable(GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem)) = False Then
PlayerMsg Index, "Your trading partner item is untradeable!", Yellow
PlayerMsg trader, "Your item is untradeable!", Yellow
Exit Sub
End If
End If

allowedLevelIndex = 10
allowedLevelTrader = 10
For i = 1 To 20
If player(Index).Bedages(i) = YES Then
allowedLevelIndex = 10 + (i * 5)
End If
Next
For i = 1 To 20
If player(trader).Bedages(i) = YES Then
allowedLevelTrader = 10 + (i * 5)
End If
Next




For i = 1 To 6
If player(trader).PokemonInstance(i).PokemonNumber > 0 Then
fs = fs + 1
End If
Next
If fs < 2 And TempPlayer(trader).TradePoke > 0 Then
PlayerMsg Index, "Your trading partner can't trade his only pokemon!", Yellow
PlayerMsg trader, "You can't trade your only pokemon!!", Yellow
Exit Sub
End If



'Player check
If TempPlayer(trader).TradePoke > 0 Then
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).level > allowedLevelTrader Then
PlayerMsg Index, "Your trading partner can't trade pokemon above level " & allowedLevelTrader, Yellow
PlayerMsg trader, "You can't trade pokemon above level " & allowedLevelTrader, Yellow
Exit Sub
End If
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).level > allowedLevelIndex Then
PlayerMsg trader, "Your trading partner can't trade pokemon above level " & allowedLevelIndex, Yellow
PlayerMsg Index, "You can't trade pokemon above level " & allowedLevelIndex, Yellow
Exit Sub
End If
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).PokemonNumber > 0 Then
fs = 0
For i = 1 To 6
If player(Index).PokemonInstance(i).PokemonNumber = 0 Then
fs = i
Exit For
End If
Next
If fs = 0 Then
PlayerMsg trader, "Your trading partner has 6 pokemons with him!", Yellow
PlayerMsg Index, "You have 6 pokemons with you!", Yellow
Exit Sub
End If
End If
End If

For i = 1 To 6
If player(Index).PokemonInstance(i).PokemonNumber > 0 Then
fs = fs + 1
End If
Next
If fs < 2 And TempPlayer(Index).TradePoke > 0 Then
PlayerMsg trader, "Your trading partner can't trade his only pokemon!", Yellow
PlayerMsg Index, "You can't trade your only pokemon!!", Yellow
Exit Sub
End If

'------------------------------------------------
'Finally we are doing trade
If TempPlayer(Index).TradePoke > 0 Then
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).level > allowedLevelTrader Then
PlayerMsg Index, "Your trading partner can't trade pokemon above level " & allowedLevelTrader, Yellow
PlayerMsg trader, "You can't trade pokemon above level " & allowedLevelTrader, Yellow
Exit Sub
End If
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).level > allowedLevelIndex Then
PlayerMsg trader, "Your trading partner can't trade pokemon above level " & allowedLevelIndex, Yellow
PlayerMsg Index, "You can't trade pokemon above level " & allowedLevelIndex, Yellow
Exit Sub
End If
If player(Index).PokemonInstance(TempPlayer(Index).TradePoke).PokemonNumber > 0 Then
fs = 0
For i = 1 To 6
If player(trader).PokemonInstance(i).PokemonNumber = 0 Then
fs = i
Exit For
End If
Next
If fs = 0 Then
PlayerMsg Index, "Your trading partner has 6 pokemons with him!", Yellow
PlayerMsg trader, "You have 6 pokemons with you!", Yellow
Exit Sub
End If
player(trader).PokemonInstance(fs) = player(Index).PokemonInstance(TempPlayer(Index).TradePoke)
player(Index).PokemonInstance(TempPlayer(Index).TradePoke).PokemonNumber = 0
SendPlayerPokemon Index
SendPlayerPokemon trader
End If
End If
'-------------------------------------------------------------------------------------------------------
If TempPlayer(trader).TradePoke > 0 Then
If player(trader).PokemonInstance(TempPlayer(trader).TradePoke).PokemonNumber > 0 Then
fs = 0
For i = 1 To 6
If player(Index).PokemonInstance(i).PokemonNumber = 0 Then
fs = i
Exit For
End If
Next
If fs = 0 Then
PlayerMsg trader, "Your trading partner has 6 pokemons with him!", Yellow
PlayerMsg Index, "You have 6 pokemons with you!", Yellow
Exit Sub
End If
player(Index).PokemonInstance(fs) = player(trader).PokemonInstance(TempPlayer(trader).TradePoke)
player(trader).PokemonInstance(TempPlayer(trader).TradePoke).PokemonNumber = 0
SendPlayerPokemon Index
SendPlayerPokemon trader
End If
End If
'-----------------------------------------------------------------------------------------------------------
If TempPlayer(Index).TradeItem > 0 Then
If GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem) > 0 Then
If item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem)).Type = ITEM_TYPE_CURRENCY Then
GiveItem trader, GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem), TempPlayer(Index).TradeItemVal
TakeItem Index, GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem), TempPlayer(Index).TradeItemVal
Else
GiveItem trader, GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem), 1
TakeItem Index, GetPlayerInvItemNum(Index, TempPlayer(Index).TradeItem), 1
End If
End If
End If
'-------------------------------------------------------------------------------------------------------------
If TempPlayer(trader).TradeItem > 0 Then
If GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem) > 0 Then
If item(GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem)).Type = ITEM_TYPE_CURRENCY Then
GiveItem Index, GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem), TempPlayer(trader).TradeItemVal
TakeItem trader, GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem), TempPlayer(trader).TradeItemVal
Else
GiveItem Index, GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem), 1
TakeItem trader, GetPlayerInvItemNum(trader, TempPlayer(trader).TradeItem), 1
End If
End If
End If

SendInventory Index
SendInventory trader
SendPokemon Index
SendPokemon trader
TempPlayer(Index).TradeName = ""
TempPlayer(Index).isTrading = NO
TempPlayer(Index).TradePoke = 0
TempPlayer(Index).TradeItem = 0
TempPlayer(Index).TradeItemVal = 0
TempPlayer(Index).TradeLocked = NO
TempPlayer(trader).TradeName = ""
TempPlayer(trader).isTrading = NO
TempPlayer(trader).TradePoke = 0
TempPlayer(trader).TradeItem = 0
TempPlayer(trader).TradeItemVal = 0
TempPlayer(trader).TradeLocked = NO
SendTradeStop Index
SendTradeStop trader
SendPlayerPokemon Index
SendPlayerPokemon trader
'DONE!
PlayerMsg Index, "Trade complete!", Yellow
PlayerMsg trader, "Trade complete!", Yellow
End Sub
Function GetPokemonMove(ByVal Pokemon As Long, ByVal move As Long) As String
On Error Resume Next
Dim moveStr As String
Dim pokeStr As String
pokeStr = Pokemon
moveStr = move
GetPokemonMove = GetMoveID(GetVar(App.Path & "\Data\Pokemon Data\" & pokeStr & ".ini", "DATA", "Move" & moveStr))
End Function
Public Function GetMapFlashlight(ByVal map As Long) As Long
On Error Resume Next
If GetVar(App.Path & "\Data\MapFL.ini", "DATA", "Map" & map) = "YES" Then
GetMapFlashlight = YES
Else
GetMapFlashlight = NO
End If
End Function


Sub CheckTravel(ByVal Index As Long, ByVal travel As Long)
On Error Resume Next
Select Case travel
Case 1
If GetItemSlot(Index, 1) > 0 Then
If GetPlayerInvItemValue(Index, GetItemSlot(Index, 1)) >= 100 Then
TakeItem Index, 1, 100
PlayerWarp Index, 6, 13, 10
PlayerMsg Index, "You came to Earth Town!", Yellow
End If
End If

Case 2
If GetItemSlot(Index, 1) > 0 Then
If player(Index).Bedages(1) = YES Then
If GetPlayerInvItemValue(Index, GetItemSlot(Index, 1)) >= 100 Then
TakeItem Index, 1, 100
PlayerWarp Index, 17, 20, 17
PlayerMsg Index, "You came to Flint Town!", Yellow
End If
Else
PlayerMsg Index, "You dont have gym 1 badge to travel there!", Yellow
End If
End If

Case 3
If GetItemSlot(Index, 1) > 0 Then
If player(Index).Bedages(2) = YES Then
If GetPlayerInvItemValue(Index, GetItemSlot(Index, 1)) >= 300 Then
TakeItem Index, 1, 300
PlayerWarp Index, 25, 24, 23
PlayerMsg Index, "You came to Naarden City!", Yellow
End If
Else
PlayerMsg Index, "You dont have gym 2 badge to travel there!", Yellow
End If
End If


Case 4
If GetItemSlot(Index, 1) > 0 Then
If player(Index).Bedages(3) = YES Then
If GetPlayerInvItemValue(Index, GetItemSlot(Index, 1)) >= 1000 Then
TakeItem Index, 1, 1000
PlayerWarp Index, 32, 18, 2
PlayerMsg Index, "You came to Lava Town!", Yellow
End If
Else
PlayerMsg Index, "You dont have gym 3 badge to travel there!", Yellow
End If
End If

End Select
End Sub

Sub UseItemOnPokemon(ByVal Index As Long, ByVal slot As Long, ByVal itemSlot As Long)
On Error Resume Next
If IsPlaying(Index) = False Then Exit Sub
Dim itemNum As Long
itemNum = GetPlayerInvItemNum(Index, itemSlot)
If itemNum < 1 Then Exit Sub

If item(itemNum).Type = ITEM_TYPE_POKEPOTION Then
Call UsePotion(Index, slot, item(itemNum).AddHP)
TakeItem Index, itemNum, 1
End If

If item(itemNum).Type = ITEM_TYPE_STONE Then
CheckPokemonStoneEvolution Index, slot, itemNum
End If

If item(itemNum).Type = ITEM_TYPE_HOLDING Then
If player(Index).PokemonInstance(slot).PokemonNumber > 0 Then
If player(Index).PokemonInstance(slot).HoldingItem > 0 Then
GiveItem Index, player(Index).PokemonInstance(slot).HoldingItem, 1
player(Index).PokemonInstance(slot).HoldingItem = itemNum
TakeItem Index, itemNum, 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & " is now holding " & Trim$(item(player(Index).PokemonInstance(slot).HoldingItem).Name) & " !", Yellow
Else
player(Index).PokemonInstance(slot).HoldingItem = itemNum
TakeItem Index, itemNum, 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & " is now holding " & Trim$(item(player(Index).PokemonInstance(slot).HoldingItem).Name) & " !", Yellow
End If
End If
End If

If item(itemNum).Type = ITEM_TYPE_SCRIPT Then
Select Case itemNum
Case 60 'Acrobatics
If CanPokeLearnTM(player(Index).PokemonInstance(slot).PokemonNumber, 469) Then
TempPlayer(Index).LearnMoveNumber = 469
TempPlayer(Index).LearnMovePokemon = slot
TempPlayer(Index).LearnMovePokemonName = Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name)
TempPlayer(Index).LearnMoveIsTM = True
SendLearnMove Index, slot, 469
Else
PlayerMsg Index, "Pokemon can't learn this TM!", BrightRed
End If
Case 61 'False Swipe

Case 62 'Scald

Case 100 'TP REMOVE
If player(Index).PokemonInstance(slot).PokemonNumber > 0 Then
SendRemoveTP Index, slot
TempPlayer(Index).isInTPRemoval = True
TakeItem Index, itemNum, 1
End If

End Select
End If

SendPlayerPokemon Index
End Sub

Sub PlayerFish(ByVal Index As Long)
On Error Resume Next
If TempPlayer(Index).CanFish = False Then
Exit Sub
End If

If HasItem(Index, 20) Then
Else
Exit Sub
End If

Dim x As Long
x = Rand(1, 100)
Dim y As Long
y = Rand(1, 50)
If x <= y Then
PlayerMsg Index, "Fishing...", Yellow
If TempPlayer(Index).PokemonBattle.PokemonNumber < 1 Then
initFishBattle Index
End If
TempPlayer(Index).CanFish = False
End If
End Sub

Sub SetPlayerMembership(ByVal Index As Long, ByVal Duration As Long)
On Error Resume Next
Dim pName As String
pName = GetPlayerName(Index)
Dim dateStart As String
dateStart = Format(Date, "m/d/yyyy")
CheckPlayerMembership Index, False
If isPlayerMember(Index) Then
Duration = GetPlayerMemberDuration(Index) + Duration
End If
Dim dur As String
dur = Duration
PutVar App.Path & "\Data\Membership.ini", pName, "MEMBER", "YES"
PutVar App.Path & "\Data\Membership.ini", pName, "DURATION", dur
PutVar App.Path & "\Data\Membership.ini", pName, "DATE", dateStart
CheckPlayerMembership Index
End Sub

Function isPlayerMember(ByVal Index As Long) As Boolean
On Error Resume Next
If GetVar(App.Path & "\Data\Membership.ini", GetPlayerName(Index), "MEMBER") = "YES" Then
isPlayerMember = True
Else
isPlayerMember = False
End If
End Function

Function GetPlayerMemberDuration(ByVal Index As Long) As Long
On Error Resume Next
If IsPlaying(Index) Then
If isPlayerMember(Index) = True Then
GetPlayerMemberDuration = Val(GetVar(App.Path & "\Data\Membership.ini", GetPlayerName(Index), "DURATION"))
End If
End If
End Function

Function GetPlayerMemberDate(ByVal Index As Long) As String
On Error Resume Next
If IsPlaying(Index) Then
If isPlayerMember(Index) = True Then
GetPlayerMemberDate = GetVar(App.Path & "\Data\Membership.ini", GetPlayerName(Index), "DATE")
End If
End If
End Function

Sub CheckPlayerMembership(ByVal Index As Long, Optional ByVal msg As Boolean = True)
On Error Resume Next
Dim dateNow As String
dateNow = Format(Date, "m/d/yyyy")
Dim Duration As Long
Dim dateStarted As String
If isPlayerMember(Index) = True Then
Duration = GetPlayerMemberDuration(Index)
dateStarted = GetPlayerMemberDate(Index)
If DateDiff("d", dateStarted, dateNow) >= Duration Then
PutVar App.Path & "\Data\Membership.ini", GetPlayerName(Index), "MEMBER", "NO"
PutVar App.Path & "\Data\Membership.ini", GetPlayerName(Index), "DURATION", "0"
PutVar App.Path & "\Data\Membership.ini", GetPlayerName(Index), "DATE", ""
If msg Then
PlayerMsg Index, "Your membership has expired!", Yellow
End If
Exit Sub
End If
If msg Then
PlayerMsg Index, "You have " & (Duration - DateDiff("d", dateStarted, dateNow)) & " days left of membership!", Yellow
End If
End If

End Sub

Function isPlayerMuted(ByVal Index As Long) As Boolean
On Error Resume Next
If Not IsPlaying(Index) Then Exit Function
If GetVar(App.Path & "\Data\Mutes.ini", "DATA", GetPlayerName(Index)) = "YES" Then
isPlayerMuted = True
Else
isPlayerMuted = False
End If
End Function

Sub MutePlayer(ByVal Index As Long)
On Error Resume Next
If IsPlaying(Index) Then
If isPlayerMuted(Index) Then
Call PutVar(App.Path & "\Data\Mutes.ini", "DATA", GetPlayerName(Index), "NO")
Else
Call PutVar(App.Path & "\Data\Mutes.ini", "DATA", GetPlayerName(Index), "YES")
End If
End If
End Sub


Sub CheckPokemonStoneEvolution(ByVal Index As Long, ByVal slot As Long, ByVal itemNum As Long)
Dim stoneName As String
stoneName = Trim$(item(itemNum).Name)
Dim pokestoneName As String
pokestoneName = Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Stone)
If pokestoneName = "Waterstone" Then
pokestoneName = "Water Stone"
End If
If pokestoneName = "Firestone" Then
pokestoneName = "Fire Stone"
End If
If pokestoneName = "Thunderstone" Then
pokestoneName = "Thunder Stone"
End If
If pokestoneName = "Leafstone" Then
pokestoneName = "Leaf Stone"
End If
If pokestoneName = "Moonstone" Then
pokestoneName = "Moon Stone"
End If
If pokestoneName = "Sunstone" Then
pokestoneName = "Sun Stone"
End If
If Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) = "Eevee" Then
Select Case stoneName
Case "Water Stone"
TempPlayer(Index).SpecialEvolveSlot = slot
TempPlayer(Index).SpecialEvolveTo = 134
TempPlayer(Index).SpecialEvolveItem = itemNum
SendEvolve Index, slot, 134
Case "Fire Stone"
TempPlayer(Index).SpecialEvolveSlot = slot
TempPlayer(Index).SpecialEvolveTo = 136
TempPlayer(Index).SpecialEvolveItem = itemNum
SendEvolve Index, slot, 136
Case "Thunder Stone"
TempPlayer(Index).SpecialEvolveSlot = slot
TempPlayer(Index).SpecialEvolveTo = 135
TempPlayer(Index).SpecialEvolveItem = itemNum
SendEvolve Index, slot, 135
End Select
Exit Sub
End If
If Trim$(item(itemNum).Name) = pokestoneName Then
'Can evolve
TempPlayer(Index).SpecialEvolveSlot = slot
TempPlayer(Index).SpecialEvolveTo = Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).EvolvesTo
TempPlayer(Index).SpecialEvolveItem = itemNum
SendEvolve Index, slot, Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).EvolvesTo
End If
End Sub

Sub CheckForLoginItem(ByVal Index As Long)
On Error Resume Next
Dim Name As String
Name = GetPlayerName(Index)
Dim items As Long
items = Val(GetVar(App.Path & "\Data\GiveOnJoin.ini", Name, "ITEMS"))
If items < 1 Then Exit Sub
Dim i As Long
For i = 1 To items
GiveItem Index, Val(GetVar(App.Path & "\Data\GiveOnJoin.ini", Name, "ITEM" & i)), Val(GetVar(App.Path & "\Data\GiveOnJoin.ini", Name, "ITEM" & i & "VAL"))
PutVar App.Path & "\Data\GiveOnJoin.ini", Name, "ITEM" & i, ""
PutVar App.Path & "\Data\GiveOnJoin.ini", Name, "ITEM" & i & "VAL", ""
Next
PutVar App.Path & "\Data\GiveOnJoin.ini", Name, "ITEMS", ""
End Sub

Function GetPlayerProfilePicture(ByVal Index As Long) As String
On Error Resume Next
GetPlayerProfilePicture = GetVar(App.Path & "\Data\alive\" & GetPlayerName(Index) & ".ini", "Other", "Profile")
End Function

Sub SetPlayerProfilePicture(ByVal Index As Long, ByVal link As String)
On Error Resume Next
Call PutVar(App.Path & "\Data\alive\" & GetPlayerName(Index) & ".ini", "Other", "Profile", link)
End Sub
Sub RemoveTP(ByVal Index As Long, ByVal slot As Long, ByVal Stat As Long)
If TempPlayer(Index).isInTPRemoval = True Then
Select Case Stat
Case STAT_HP
player(Index).PokemonInstance(slot).MaxHp = player(Index).PokemonInstance(slot).MaxHp - 2
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s HP has been reduced by 2! +1 TP", Yellow
Case STAT_ATK
player(Index).PokemonInstance(slot).atk = player(Index).PokemonInstance(slot).atk - 1
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s ATK has been reduced by 1! +1 TP", Yellow
Case STAT_DEF
player(Index).PokemonInstance(slot).def = player(Index).PokemonInstance(slot).def - 1
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s DEF has been reduced by 1! +1 TP", Yellow
Case STAT_SPATK
player(Index).PokemonInstance(slot).spatk = player(Index).PokemonInstance(slot).spatk - 1
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s SPATK has been reduced by 1! +1 TP", Yellow
Case STAT_SPDEF
player(Index).PokemonInstance(slot).spdef = player(Index).PokemonInstance(slot).spdef - 1
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s SPDEF has been reduced by 1! +1 TP", Yellow
Case STAT_SPEED
player(Index).PokemonInstance(slot).spd = player(Index).PokemonInstance(slot).spd - 1
player(Index).PokemonInstance(slot).TP = player(Index).PokemonInstance(slot).TP + 1
PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(slot).PokemonNumber).Name) & "'s SPEED has been reduced by 1! +1 TP", Yellow
End Select
SendPlayerPokemon Index
SendPlayerData Index
TempPlayer(Index).isInTPRemoval = False
End If
End Sub







Public Sub AddRankedPoint(ByVal Index As Long)
On Error Resume Next
Dim rp As Long
rp = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "RankPoints"))
Dim newrp As Long
newrp = rp + 1
Dim str As String
str = newrp
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "RankPoints", Trim$(newrp))
End Sub



Public Sub RemoveRankedPoint(ByVal Index As Long)
On Error Resume Next
Dim rp As Long
rp = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "RankPoints"))
Dim newrp As Long
newrp = rp - 1
Dim str As String
str = newrp
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "RankPoints", Trim$(newrp))
End Sub
Function GetRankedPoints(ByVal Index As Long) As Long
On Error Resume Next
GetRankedPoints = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "RankPoints"))
End Function

Public Function GetPlayerDivision(ByVal Index As Long, ByVal RankPoints As Long) As Long
On Error Resume Next
If RankPoints < 5 Then
     GetPlayerDivision = DIVISION_BRONZE_3
     End If
     If RankPoints >= 5 And RankPoints < 15 Then
     GetPlayerDivision = DIVISION_BRONZE_2
     End If
      If RankPoints >= 15 And RankPoints < 25 Then
     GetPlayerDivision = DIVISION_BRONZE_1
     End If
     If RankPoints >= 25 And RankPoints < 35 Then
     GetPlayerDivision = DIVISION_SILVER_3
     End If
      If RankPoints >= 35 And RankPoints < 45 Then
     GetPlayerDivision = DIVISION_SILVER_2
     End If
      If RankPoints >= 45 And RankPoints < 55 Then
     GetPlayerDivision = DIVISION_SILVER_1
     End If
      If RankPoints >= 55 And RankPoints < 65 Then
     GetPlayerDivision = DIVISION_GOLD_3
     End If
     If RankPoints >= 65 And RankPoints < 75 Then
     GetPlayerDivision = DIVISION_GOLD_2
     End If
     If RankPoints >= 75 And RankPoints < 90 Then
     GetPlayerDivision = DIVISION_GOLD_1
     End If
     If RankPoints >= 90 And RankPoints < 105 Then
     GetPlayerDivision = DIVISION_PLATINUM_3
     End If
     If RankPoints >= 105 And RankPoints < 120 Then
     GetPlayerDivision = DIVISION_PLATINUM_2
     End If
     If RankPoints >= 120 And RankPoints < 150 Then
     GetPlayerDivision = DIVISION_PLATINUM_1
     End If
     If RankPoints >= 150 And RankPoints < 175 Then
     GetPlayerDivision = DIVISION_DIAMOND_3
     End If
      If RankPoints >= 175 And RankPoints < 200 Then
     GetPlayerDivision = DIVISION_DIAMOND_2
     End If
      If RankPoints >= 200 Then
     GetPlayerDivision = DIVISION_DIAMOND_1
     End If
End Function
Function GetPlayerJournal(ByVal Index As Long) As String
GetPlayerJournal = ReadText("Data\journals\" & GetPlayerName(Index) & ".txt")
End Function

Function GetPlayerPlaytimeMinutes(ByVal Index As Long) As Long
GetPlayerPlaytimeMinutes = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Minutes"))
End Function
Function GetPlayerPlaytimeHours(ByVal Index As Long) As Long
GetPlayerPlaytimeHours = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Hours"))
End Function

Sub AddMinutePlaytime(ByVal Index As Long)
Dim mins As Long
Dim hours As Long
mins = GetPlayerPlaytimeMinutes(Index)
hours = GetPlayerPlaytimeHours(Index)
mins = mins + 1
If mins >= 60 Then
mins = 0
hours = hours + 1
End If
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Hours", Trim$(hours))
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Minutes", Trim$(mins))
End Sub

Sub AddEgg(ByVal Index As Long)
TempPlayer(Index).eggExpTemp = 0
TempPlayer(Index).eggStepsTemp = 0
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Egg", "YES")
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggSteps", "50000")
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggEXP", "100000")
End Sub

Sub RemoveEgg(ByVal Index As Long)
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Egg", "NO")
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggSteps", "0")
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggEXP", "0")
End Sub

Function DoesPlayerHaveEgg(ByVal Index As Long) As Boolean
If GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Egg") = "YES" Then
DoesPlayerHaveEgg = True
End If
End Function

Sub UpdatePlayerEgg(ByVal Index As Long, ByVal reduceSteps As Long, ByVal reduceEXP As Long)
If DoesPlayerHaveEgg(Index) = False Then Exit Sub
Dim expAmount As Long
Dim stepsAmount As Long
stepsAmount = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggSteps"))
expAmount = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggEXP"))
stepsAmount = stepsAmount - reduceSteps
expAmount = expAmount - reduceEXP
If stepsAmount <= 0 Then
stepsAmount = 0
End If
If expAmount <= 0 Then
expAmount = 0
End If
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggSteps", Trim$(stepsAmount))
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggEXP", Trim$(expAmount))
End Sub

Function GetPlayerEggEXP(ByVal Index As Long) As Long
GetPlayerEggEXP = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggEXP"))
End Function
Function GetPlayerEggSteps(ByVal Index As Long) As Long
GetPlayerEggSteps = Val(GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "EggSteps"))
End Function

Sub SaveEggFromTemp(ByVal Index As Long)
If DoesPlayerHaveEgg(Index) Then
UpdatePlayerEgg Index, TempPlayer(Index).eggStepsTemp, TempPlayer(Index).eggExpTemp
TempPlayer(Index).eggExpTemp = 0
TempPlayer(Index).eggStepsTemp = 0
End If
End Sub
'HATCH
Sub HatchEgg(ByVal Index As Long)
If DoesPlayerHaveEgg(Index) = False Then Exit Sub
Dim pokeNum As Long
If GetPlayerEggSteps(Index) <= 0 And GetPlayerEggEXP(Index) <= 0 Then
pokeNum = GetEggPokemon
GivePokemon Index, pokeNum, 1, 0, 1, 0, 10
SendDialog Index, "Congratulations " & Trim$(GetPlayerName(Index)) & ", baby " & Trim$(Pokemon(pokeNum).Name) & " has just hatched from egg!", 1
SendDialog Index, "You should be proud!", 1
'PlayerMsg index, "Congratulations, baby " & Trim$(Pokemon(pokeNum).Name) & " has just hatched from egg!", Yellow
RemoveEgg Index
End If
End Sub

Function GetEggPokemon() As Long
Dim eggNums As Long
eggNums = Val(GetVar(App.Path & "\Data\EggHatch.ini", "DATA", "Num"))
Dim pokeNum As Long
Dim rndNum As Long
rndNum = Rand(1, eggNums)
pokeNum = Val(GetVar(App.Path & "\Data\EggHatch.ini", "DATA", Trim$(rndNum)))
GetEggPokemon = pokeNum
End Function
Public Sub UseBike(ByVal Index As Long)
Select Case TempPlayer(Index).HasBike

Case YES
TempPlayer(Index).HasBike = NO
player(Index).Sprite = 509
SendPlayerData Index
Case NO
TempPlayer(Index).HasBike = YES
player(Index).Sprite = 514
SendPlayerData Index
End Select
End Sub


Function DoesPlayerHaveBike(ByVal Index As Long) As Boolean
On Error Resume Next
Dim str As String
str = GetVar(App.Path & "\Data\alive\" & GetPlayerName(Index) & ".ini", "Other", "Bike")
If Trim$(str) = "YES" Then
DoesPlayerHaveBike = True
End If
End Function
