Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
On Error Resume Next
    'frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next

End Sub

Sub SendMusicToOne(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SMapMusic
buffer.WriteString GetVar(App.Path & "\Data\MapData\" & GetPlayerMap(Index) & ".ini", "DATA", "Music")
SendDataTo Index, buffer.ToArray()
Set buffer = Nothing
End Sub




Sub SendMusicToMap(ByVal map As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SMapMusic
buffer.WriteString GetVar(App.Path & "\Data\MapData\" & map & ".ini", "DATA", "Music")
SendDataToMap map, buffer.ToArray()
Set buffer = Nothing
End Sub

Sub SendIntro(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SIntro
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendIsInBattle(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SisInBattle
buffer.WriteLong Index
If TempPlayer(Index).PokemonBattle.PokemonNumber > 0 And TempPlayer(Index).PokemonBattle.hp > 0 Then
buffer.WriteLong YES
Else
buffer.WriteLong NO
End If
SendDataToAll buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendBattleInfo(ByVal Index As Long, ByVal pokecoins As Long, ByVal win As Long, ByVal EXP As Long)
On Error Resume Next
SendIsInBattle (Index)
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SBattleInfo
If TempPlayer(Index).PokemonBattle.MapSlot < 1 Then
buffer.WriteLong 1
Else
buffer.WriteLong map(player(Index).map).Pokemon(TempPlayer(Index).PokemonBattle.MapSlot).Chance
End If
If TempPlayer(Index).PokemonBattle.PokemonNumber < 1 Then
Else
buffer.WriteString Pokemon(TempPlayer(Index).PokemonBattle.PokemonNumber).Name
End If
buffer.WriteLong pokecoins
buffer.WriteLong win
buffer.WriteLong EXP
SendDataTo Index, buffer.ToArray
Set buffer = Nothing


End Sub

Sub SendOpenSwitch(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SOpenSwitch
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub


Sub SendPCScanRequest(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SPCRequest
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
On Error Resume Next
    If frmServer.Socket(Index).state = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean
On Error Resume Next
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal Index As Long) As Boolean
On Error Resume Next
    If IsConnected(Index) Then
        If LenB(Trim$(player(Index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsConnected(i) Then
            If LCase$(Trim$(player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
On Error Resume Next
    Dim i As Long
    Dim n As Long

    For i = 1 To MAX_PLAYERS

        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
On Error Resume Next
    Dim FileName As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    FileName = App.Path & "\data\banlist.ini"
        IsBanned = False
    ' Check if file exists
    If GetVar(FileName, "DATA", IP) = "YES" Then
    IsBanned = True
    End If
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    
    If IsConnected(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        buffer.WriteBytes Data()
        frmServer.Socket(Index).SendData buffer.ToArray()
        DoEvents
        Set buffer = Nothing
    End If

End Sub

Sub SendDataToAll(ByRef Data() As Byte)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If i <> Index Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByRef Data() As Byte)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal mapnum As Long, ByRef Data() As Byte)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                If i <> Index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If

    Next

End Sub

Public Sub GlobalMsg(ByVal msg As String, ByVal color As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SGlobalMsg
    buffer.WriteString msg
    buffer.WriteLong color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal msg As String, ByVal color As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SAdminMsg
    buffer.WriteString msg
    buffer.WriteLong color

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal msg As String, ByVal color As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerMsg
    buffer.WriteString msg
    buffer.WriteLong color
    SendDataTo Index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal msg As String, ByVal color As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteInteger SMapMsg
    buffer.WriteString msg
    buffer.WriteLong color
    SendDataToMap mapnum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub VersionCheck(ByVal Index As Long, ByVal version As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteInteger SVersionCheck
    buffer.WriteLong version
    SendDataTo Index, buffer.ToArray
    Set buffer = Nothing
End Sub


Public Sub TotalPlayersCheck(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteInteger STotalPlayersCheck
    buffer.WriteLong TotalPlayersOnline
    SendDataTo Index, buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub AdminCheck(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SAdminCheck
    If AdminOnly = True Then
    buffer.WriteLong YES
    Else
    buffer.WriteLong NO
    End If
    SendDataTo Index, buffer.ToArray
    Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal msg As String)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteInteger SAlertMsg
    buffer.WriteString msg
    SendDataTo Index, buffer.ToArray
    NewDoEvents
    Call CloseSocket(Index)
    
    Set buffer = Nothing
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal reason As String)
On Error Resume Next
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call GlobalMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & reason & ")", White)
        End If

        Call AlertMsg(Index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
On Error Resume Next
    Dim i As Long

    If (Index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

End Sub

Sub SocketConnected(ByVal Index As Long)
On Error Resume Next
    If Index <> 0 Then

        ' Are they trying to connect more then one connection?
        'If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
        If Not IsBanned(GetPlayerIP(Index)) Then
        
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
            Call VersionCheck(Index, VersionCode)
            Call TotalPlayersCheck(Index)
        Else
            Call AlertMsg(Index, "You have been banned from PEO , and can no longer play.")
        End If

        'Else
        ' Tried multiple connections
        '    Call AlertMsg(Index, Options.Game_Name & " does not allow multiple IP's anymore.")
        'End If
    End If

End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
On Error Resume Next
    Dim buffer() As Byte
    Dim Data() As Byte
    Dim pLength As Long
    Dim tcpValue As Long
    frmServer.Socket(Index).GetData buffer(), vbUnicode, DataLength
    'IncomingBytes = IncomingBytes + DataLength
    TempPlayer(Index).buffer.WriteBytes buffer()

    'TempPlayer(Index).Buffer.DecompressBuffer
    If TempPlayer(Index).buffer.Length >= 4 Then
        pLength = TempPlayer(Index).buffer.ReadLong(False)

        If pLength < 0 Then
            Exit Sub
        End If
    End If

    Do While pLength > 0 And pLength <= TempPlayer(Index).buffer.Length - 4

        If pLength <= TempPlayer(Index).buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).buffer.ReadLong
            Data() = TempPlayer(Index).buffer.ReadBytes(pLength + 1)
            
            
            'If EncryptPackets Then
            '    Encryption_XOR_DecryptByte Data(), PacketKeys(Player(Index).PacketInIndex)
            '    Player(Index).PacketInIndex = Player(Index).PacketInIndex + 1
            '    If Player(Index).PacketInIndex > PacketEncKeys - 1 Then Player(Index).PacketInIndex = 0
            'End If
            HandleData Index, Data()
        End If

        pLength = 0

        If TempPlayer(Index).buffer.Length >= 4 Then
            pLength = TempPlayer(Index).buffer.ReadLong(False)

            If pLength < 0 Then
                Exit Sub
            End If
        End If

    Loop

    If GetPlayerAccess(Index) <= 0 Then

        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then
            HackingAttempt Index, "Data Flooding"
            Exit Sub
        End If

        ' Check for packet flooding
        If TempPlayer(Index).DataPackets > 500 Then
            HackingAttempt Index, "Packet Flooding"
            Exit Sub
        End If
    End If

    ' Check if elapsed time has passed
    'Player(Index).DataBytes = Player(Index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = GetTickCount + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If

    If TempPlayer(Index).buffer.Length <= 1 Then TempPlayer(Index).buffer.Flush
End Sub

Sub CloseSocket(ByVal Index As Long)
On Error Resume Next
    If Index > 0 Then
        Call LeftGame(Index)
        Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        frmServer.Socket(Index).Close
        Call UpdateCaption
        Call ClearPlayer(Index)
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
On Error Resume Next
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong mapnum
    buffer.WriteString Trim$(map(mapnum).Name)
    buffer.WriteLong map(mapnum).Revision
    buffer.WriteLong map(mapnum).Moral
    buffer.WriteLong map(mapnum).Tileset
    buffer.WriteLong map(mapnum).Up
    buffer.WriteLong map(mapnum).Down
    buffer.WriteLong map(mapnum).Left
    buffer.WriteLong map(mapnum).Right
    buffer.WriteLong map(mapnum).Music
    buffer.WriteLong map(mapnum).BootMap
    buffer.WriteLong map(mapnum).BootX
    buffer.WriteLong map(mapnum).BootY
    buffer.WriteLong map(mapnum).MaxX
    buffer.WriteLong map(mapnum).MaxY

    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY

            With map(mapnum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteByte .Layer(i).x
                    buffer.WriteByte .Layer(i).y
                    buffer.WriteByte .Layer(i).Tileset
                Next
                buffer.WriteLong .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong map(mapnum).NPC(x)
    Next

    For x = 1 To MAX_MAP_POKEMONS
    buffer.WriteLong map(mapnum).Pokemon(x).PokemonNumber
    buffer.WriteLong map(mapnum).Pokemon(x).LevelFrom
    buffer.WriteLong map(mapnum).Pokemon(x).LevelTo
    buffer.WriteLong map(mapnum).Pokemon(x).Custom
    buffer.WriteLong map(mapnum).Pokemon(x).atk
    buffer.WriteLong map(mapnum).Pokemon(x).def
    buffer.WriteLong map(mapnum).Pokemon(x).spatk
    buffer.WriteLong map(mapnum).Pokemon(x).spdef
    buffer.WriteLong map(mapnum).Pokemon(x).spd
    buffer.WriteLong map(mapnum).Pokemon(x).hp
    buffer.WriteLong map(mapnum).Pokemon(x).Chance
    Next
    buffer.CompressBuffer
    MapCache(mapnum).Data = buffer.ToArray()
    
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
On Error Resume Next
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If i <> Index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        If player(Index).Access >= 1 Then
        s = "There are " & n & " other players online: " & s & "."
        Else
        s = "There are " & n & " other players online."
        End If
    End If

    Call PlayerMsg(Index, s, White)
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
On Error Resume Next
    Dim buffer As clsBuffer, i As Long, a As Long

    If Index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerData
    buffer.WriteLong Index
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerLevel(Index)
    buffer.WriteLong GetPlayerPOINTS(Index)
    buffer.WriteLong GetPlayerSprite(Index)
    buffer.WriteLong GetPlayerMap(Index)
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteLong player(Index).mood
    If TempPlayer(Index).notVisible Then
    buffer.WriteLong YES
    Else
    buffer.WriteLong NO
    End If
    For i = 1 To 6
    buffer.WriteLong player(Index).PokemonInstance(i).PokemonNumber
    Next
    For i = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(Index, i)
    Next
    
    For a = 1 To MAX_GYMS
    buffer.WriteLong player(Index).Bedages(a)
    Next
    buffer.WriteLong player(Index).Equipment(Equipment.Armor)
    buffer.WriteLong player(Index).Equipment(Equipment.Helmet)
    buffer.WriteLong player(Index).Equipment(Equipment.Shield)
    buffer.WriteLong player(Index).Equipment(Equipment.Weapon)
    buffer.WriteLong player(Index).Equipment(Equipment.Mask)
    buffer.WriteLong player(Index).Equipment(Equipment.Outfit)
    buffer.WriteLong TempPlayer(Index).HasBike
    If isPlayerMember(Index) = True Then
    buffer.WriteLong YES
    Else
    buffer.WriteLong YES
    End If
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SLeft
    buffer.WriteLong Index
    SendDataToMapBut Index, mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
On Error Resume Next
    'SendDataToMap GetPlayerMap(Index),PlayerData(Index)
    SendDataToAll PlayerData(Index)
    SendMapEquipment Index
End Sub

Sub SendMap(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(mapnum).Data) - LBound(MapCache(mapnum).Data)) + 5
    buffer.WriteInteger SMapData
    buffer.WriteBytes MapCache(mapnum).Data()
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
    
End Sub

Sub SendCustomMap(ByVal Index As Long, ByVal mapnum As Long, ByVal mapDir As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.PreAllocate (UBound(MapCache(mapnum).Data) - LBound(MapCache(mapnum).Data)) + 5
buffer.WriteInteger SMapData
buffer.WriteBytes MapCache(mapnum).Data()
buffer.WriteLong mapDir
SendDataTo Index, buffer.ToArray()
Set buffer = Nothing

End Sub


Sub SendTrainerCard(ByVal Index As Long, ByVal playerIndex As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim i As Long
Set buffer = New clsBuffer
buffer.WriteInteger STrainerCard
buffer.WriteString player(playerIndex).Name
For i = 1 To 6
buffer.WriteLong player(playerIndex).PokemonInstance(i).PokemonNumber
buffer.WriteLong player(playerIndex).PokemonInstance(i).level
buffer.WriteLong player(playerIndex).PokemonInstance(i).isShiny
buffer.WriteLong player(playerIndex).Bedages(i)
Next
buffer.WriteString GetPlayerProfilePicture(playerIndex)
buffer.WriteLong GetRankedPoints(playerIndex)
If DoesCrewExist(GetPlayerCrew(playerIndex)) Then
buffer.WriteString GetPlayerCrew(playerIndex)
buffer.WriteString GetCrewPicture(GetPlayerCrew(playerIndex))
Else
buffer.WriteString "None"
buffer.WriteString ""
End If
If GetPlayerCrew(Index) <> "" Then
buffer.WriteLong YES
Else
buffer.WriteLong NO
End If

SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendOpenBank(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SOpenBank
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub



Sub SendUpdateBank(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SUpdateBank
buffer.WriteLong GetPlayerInvItemValue(Index, 1)
buffer.WriteLong player(Index).StoredPC
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub


Sub SendOpenStorage(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SOpenStorage
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteLong MapItem(mapnum, i).Num
        buffer.WriteLong MapItem(mapnum, i).value
        buffer.WriteLong MapItem(mapnum, i).x
        buffer.WriteLong MapItem(mapnum, i).y
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapnum As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteLong MapItem(mapnum, i).Num
        buffer.WriteLong MapItem(mapnum, i).value
        buffer.WriteLong MapItem(mapnum, i).x
        buffer.WriteLong MapItem(mapnum, i).y
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal MapNpcNum As Byte)
On Error Resume Next
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapNpcVitals
    buffer.WriteByte MapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(mapnum).NPC(MapNpcNum).Vital(i)
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(i).Num
        buffer.WriteLong MapNpc(mapnum).NPC(i).x
        buffer.WriteLong MapNpc(mapnum).NPC(i).y
        buffer.WriteLong MapNpc(mapnum).NPC(i).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(i).Vital(hp)
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(i).Num
        buffer.WriteLong MapNpc(mapnum).NPC(i).x
        buffer.WriteLong MapNpc(mapnum).NPC(i).y
        buffer.WriteLong MapNpc(mapnum).NPC(i).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(i).Vital(hp)
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendItems(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If LenB(Trim$(item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If

    Next

End Sub

Sub SendAnimations(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If

    Next

End Sub

Sub SendNpcs(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_NPCS

        If LenB(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If

    Next

End Sub

Sub SendResources(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If

    Next

End Sub

Sub SendPokemon(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_POKEMONS

        If LenB(Trim$(Pokemon(i).Name)) > 0 Then
            Call SendUpdatePokemonTo(Index, i)
        End If

    Next

End Sub

Sub SendMove(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MOVES

        If LenB(Trim$(PokemonMove(i).Name)) > 0 Then
            Call SendUpdateMoveTo(Index, i)
        End If

    Next

End Sub

Sub SendInventory(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(Index, i)
        buffer.WriteLong GetPlayerInvItemValue(Index, i)
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStoragePokemonLoad(ByVal Index As Long, ByVal slot As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SStorageLoadPoke
buffer.WriteLong slot
Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal invslot As Long)
On Error Resume Next
SendInventory Index
    'Dim packet As String
    'Dim buffer As clsBuffer
    'Set buffer = New clsBuffer
    
    'buffer.WriteInteger SPlayerInvUpdate
    'buffer.WriteLong invslot
    'buffer.WriteLong GetPlayerInvItemNum(index, invslot)
    'buffer.WriteLong GetPlayerInvItemValue(index, invslot)
    'SendDataTo index, buffer.ToArray()
    
    'Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapWornEq
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerEquipment(Index, Armor)
    buffer.WriteLong GetPlayerEquipment(Index, Weapon)
    buffer.WriteLong GetPlayerEquipment(Index, Helmet)
    buffer.WriteLong GetPlayerEquipment(Index, Shield)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal playerNum As Long, ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SMapWornEq
    buffer.WriteLong playerNum
    buffer.WriteLong GetPlayerEquipment(playerNum, Armor)
    buffer.WriteLong GetPlayerEquipment(playerNum, Weapon)
    buffer.WriteLong GetPlayerEquipment(playerNum, Helmet)
    buffer.WriteLong GetPlayerEquipment(playerNum, Shield)
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case hp
            buffer.WriteInteger SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.hp)
            buffer.WriteLong GetPlayerVital(Index, Vitals.hp)
        Case mp
            buffer.WriteInteger SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.mp)
            buffer.WriteLong GetPlayerVital(Index, Vitals.mp)
        Case SP
            'Buffer.WriteLong SPlayerSp
            'Buffer.WriteLong GetPlayerMaxVital(index, Vitals.SP)
            'Buffer.WriteLong GetPlayerVital(index, Vitals.SP)
    End Select

    'SendDataTo index, Buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendEXP(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerEXP
    buffer.WriteLong GetPlayerExp(Index)
    buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStats(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SPlayerStats
    buffer.WriteLong GetPlayerStat(Index, Stats.strength)
    buffer.WriteLong GetPlayerStat(Index, Stats.endurance)
    buffer.WriteLong GetPlayerStat(Index, Stats.vitality)
    buffer.WriteLong GetPlayerStat(Index, Stats.willpower)
    buffer.WriteLong GetPlayerStat(Index, Stats.intelligence)
    buffer.WriteLong GetPlayerStat(Index, Stats.spirit)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)
On Error Resume Next
    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(Index, "Welcome to Poketopia! Please report all bugs on the discord.", White)
        Call PlayerMsg(Index, "Official Discord: https://discord.gg/qWrEU2t!", White)
        Call PlayerMsg(Index, "Latest Updates:", Yellow)
        'Call PlayerMsg(index, "2017", White)
        'Call PlayerMsg(index, "News: Generation 6", White)
    End If

    ' Send whos online
    'Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SClassesData
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.hp)
        buffer.WriteLong GetClassMaxVital(i, Vitals.mp)
        buffer.WriteLong GetClassMaxVital(i, Vitals.SP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        buffer.WriteLong Class(i).Stat(Stats.strength)
        buffer.WriteLong Class(i).Stat(Stats.endurance)
        buffer.WriteLong Class(i).Stat(Stats.vitality)
        buffer.WriteLong Class(i).Stat(Stats.intelligence)
        buffer.WriteLong Class(i).Stat(Stats.willpower)
        buffer.WriteLong Class(i).Stat(Stats.spirit)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long, n As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SNewCharClasses
    buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        buffer.WriteString GetClassName(i)
        buffer.WriteLong GetClassMaxVital(i, Vitals.hp)
        buffer.WriteLong GetClassMaxVital(i, Vitals.mp)
        buffer.WriteLong GetClassMaxVital(i, Vitals.SP)
        
        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        buffer.WriteLong Class(i).Stat(Stats.strength)
        buffer.WriteLong Class(i).Stat(Stats.endurance)
        buffer.WriteLong Class(i).Stat(Stats.vitality)
        buffer.WriteLong Class(i).Stat(Stats.willpower)
        buffer.WriteLong Class(i).Stat(Stats.intelligence)
        buffer.WriteLong Class(i).Stat(Stats.spirit)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLeftGame(ByVal Index As Long)
Dim i As Long
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SPlayerData
    buffer.WriteLong Index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    ''''
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    For i = 1 To 6
    buffer.WriteLong 0
    Next
    For i = 1 To Stats.Stat_Count - 1
    buffer.WriteLong 0
    Next
    For i = 1 To MAX_GYMS
    buffer.WriteLong 0
    Next
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SPlayerXY
    buffer.WriteLong GetPlayerX(Index)
    buffer.WriteLong GetPlayerY(Index)
    buffer.WriteLong GetPlayerDir(Index)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
     buffer.WriteInteger SUpdateItem
    ItemSize = LenB(item(itemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(item(itemNum)), ItemSize
    
   
    buffer.WriteLong itemNum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal itemNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteInteger SUpdateItem
    ItemSize = LenB(item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(item(itemNum)), ItemSize
    
    buffer.WriteLong itemNum
    buffer.WriteBytes ItemData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteInteger SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteInteger SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NpcNum)), NPCSize
    buffer.WriteInteger SUpdateNpc
    buffer.WriteLong NpcNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal NpcNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NpcNum)), NPCSize
    buffer.WriteInteger SUpdateNpc
    buffer.WriteLong NpcNum
    buffer.WriteBytes NPCData
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteInteger SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteInteger SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdatePokemonToAll(ByVal pokemonnum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim Pokemonize As Long
    Dim PokemonData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SUpdatePokemon
    Pokemonize = LenB(Pokemon(pokemonnum))
    ReDim PokemonData(Pokemonize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(pokemonnum)), Pokemonize
    buffer.WriteLong pokemonnum
    buffer.WriteBytes PokemonData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoveToAll(ByVal moveNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim Movesize As Long
    Dim moveData() As Byte
    
    Set buffer = New clsBuffer
    
    Movesize = LenB(PokemonMove(moveNum))
    ReDim moveData(Movesize - 1)
    CopyMemory moveData(0), ByVal VarPtr(PokemonMove(moveNum)), Movesize
    
    buffer.WriteInteger SUpdateMove
    buffer.WriteLong moveNum
    buffer.WriteBytes moveData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdatePokemonTo(ByVal Index As Long, ByVal pokemonnum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim Pokemonize As Long
    Dim PokemonData() As Byte
    
    Set buffer = New clsBuffer
    
    Pokemonize = LenB(Pokemon(pokemonnum))
    ReDim PokemonData(Pokemonize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(pokemonnum)), Pokemonize
    
    buffer.WriteInteger SUpdatePokemon
    buffer.WriteLong pokemonnum
    buffer.WriteBytes PokemonData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoveTo(ByVal Index As Long, ByVal moveNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim Movesize As Long
    Dim moveData() As Byte
    
    Set buffer = New clsBuffer
    
    Movesize = LenB(PokemonMove(moveNum))
    ReDim moveData(Movesize - 1)
    CopyMemory moveData(0), ByVal VarPtr(PokemonMove(moveNum)), Movesize
    
    buffer.WriteInteger SUpdateMove
    buffer.WriteLong moveNum
    buffer.WriteBytes moveData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendMovesToPlayer(ByVal Index As Long)
On Error Resume Next
Dim i As Long
For i = 1 To MAX_MOVES
SendUpdateMoveTo Index, i
Next
End Sub
Sub SendShops(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteInteger SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteInteger SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
        End If

    Next

End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteInteger SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal spellnum As Long)
On Error Resume Next
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteInteger SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub sendopenroster(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SOpenRoster
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendPlayerPokemon(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    Dim x As Long
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SPlayerPokemon
    For i = 1 To 6
        buffer.WriteLong player(Index).PokemonInstance(i).PokemonNumber
        buffer.WriteLong player(Index).PokemonInstance(i).level
        buffer.WriteLong player(Index).PokemonInstance(i).hp
        buffer.WriteLong player(Index).PokemonInstance(i).MaxHp
        buffer.WriteLong player(Index).PokemonInstance(i).pp
        buffer.WriteLong player(Index).PokemonInstance(i).EXP
        buffer.WriteLong player(Index).PokemonInstance(i).TP
        buffer.WriteLong player(Index).PokemonInstance(i).atk
        buffer.WriteLong player(Index).PokemonInstance(i).def
        buffer.WriteLong player(Index).PokemonInstance(i).spatk
        buffer.WriteLong player(Index).PokemonInstance(i).spdef
        buffer.WriteLong player(Index).PokemonInstance(i).spd
        buffer.WriteLong player(Index).PokemonInstance(i).isShiny
        buffer.WriteLong player(Index).PokemonInstance(i).HoldingItem
        For x = 1 To 4
        buffer.WriteLong player(Index).PokemonInstance(i).moves(x).number
        buffer.WriteLong player(Index).PokemonInstance(i).moves(x).pp
        Next
        buffer.WriteLong player(Index).PokemonInstance(i).nature
        If Not player(Index).PokemonInstance(i).level = 100 Then
        buffer.WriteLong PokemonEXP(player(Index).PokemonInstance(i).level + 1)
        Else
        buffer.WriteLong 100
        End If
        
    Next
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStorage(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteInteger SStorageUpdate
    For i = 1 To 250
        buffer.WriteLong player(Index).StoragePokemonInstance(i).PokemonNumber
        buffer.WriteLong player(Index).StoragePokemonInstance(i).level
        buffer.WriteLong player(Index).StoragePokemonInstance(i).nature
        buffer.WriteLong player(Index).StoragePokemonInstance(i).atk
        buffer.WriteLong player(Index).StoragePokemonInstance(i).def
        buffer.WriteLong player(Index).StoragePokemonInstance(i).spd
        buffer.WriteLong player(Index).StoragePokemonInstance(i).spatk
        buffer.WriteLong player(Index).StoragePokemonInstance(i).spdef
        buffer.WriteLong player(Index).StoragePokemonInstance(i).MaxHp
        buffer.WriteLong player(Index).StoragePokemonInstance(i).isShiny
    Next
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
On Error Resume Next
    Dim packet As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(Index, i)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_num As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteInteger SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then

        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).x
            buffer.WriteLong ResourceCache(GetPlayerMap(Index)).ResourceData(i).y
        Next

    End If

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapnum As Long, ByVal Resource_num As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteInteger SResourceCache
    buffer.WriteLong ResourceCache(mapnum).Resource_Count

    If ResourceCache(mapnum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(mapnum).Resource_Count
            buffer.WriteByte ResourceCache(mapnum).ResourceData(i).ResourceState
            buffer.WriteLong ResourceCache(mapnum).ResourceData(i).x
            buffer.WriteLong ResourceCache(mapnum).ResourceData(i).y
        Next

    End If

    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SDoorAnimation
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapnum As Long, ByVal message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SActionMsg
    buffer.WriteString message
    buffer.WriteLong color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong Len(message)
    buffer.WriteLong 16
    buffer.WriteLong DarkGrey
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap mapnum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub



Sub SendAnimation(ByVal mapnum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendSound(ByVal Index As Long, sound As String)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SSound
    buffer.WriteString sound
    SendDataTo Index, buffer.ToArray
    Set buffer = Nothing
End Sub

Sub SendCooldown(ByVal Index As Long, ByVal slot As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SCooldown
    buffer.WriteLong slot
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub



Sub SayMsg_Map(ByVal mapnum As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString message
    buffer.WriteString "[Map]"
    buffer.WriteLong saycolour
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString message
    buffer.WriteString "[Global]"
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SResetShopAction
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SStunned
    buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendNpcBattle(ByVal Index As Long, ByVal myslot As Long, Optional ByVal isNPCBattle As Long = NO, Optional ByVal NPCplayMusic As Long = NO, Optional ByVal NPCMusic As String = "Battle1.mp3", Optional ByVal customBackground As String = "")
    On Error Resume Next
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    'Send enemy pokemons
    buffer.WriteInteger SNpcBattle
    buffer.WriteLong TempPlayer(Index).PokemonBattle.PokemonNumber
    buffer.WriteLong TempPlayer(Index).PokemonBattle.level
    buffer.WriteLong TempPlayer(Index).PokemonBattle.hp
    buffer.WriteLong TempPlayer(Index).PokemonBattle.MaxHp
    buffer.WriteLong TempPlayer(Index).PokemonBattle.isShiny
    'Send my Pokemons
    buffer.WriteLong myslot
    buffer.WriteLong TempPlayer(Index).BattleCurrentTurn
    buffer.WriteLong isNPCBattle
    buffer.WriteLong NPCplayMusic
    buffer.WriteString NPCMusic
    buffer.WriteString customBackground
    'Before everything send pokemon data to player
    SendPlayerPokemon Index
    
    'Send start battle
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
    SendIsInBattle Index
End Sub

Sub SendBattleUpdate(ByVal Index As Long, ByVal myslot As Long, Optional ByVal UnblockBattle As Long = 0, Optional ByVal PVPUnblock As Long = 0)
On Error Resume Next
Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SBattleUpdate
    buffer.WriteLong TempPlayer(Index).PokemonBattle.PokemonNumber
    buffer.WriteLong TempPlayer(Index).PokemonBattle.hp
    buffer.WriteLong TempPlayer(Index).PokemonBattle.MaxHp
    buffer.WriteLong TempPlayer(Index).PokemonBattle.level
    'shiny *.*
    buffer.WriteLong TempPlayer(Index).PokemonBattle.isShiny
    buffer.WriteLong UnblockBattle
    buffer.WriteLong PVPUnblock
    'My pokemon
    buffer.WriteLong myslot
    buffer.WriteLong TempPlayer(Index).BattleCurrentTurn
    'And send all
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing

    SendIsInBattle Index

End Sub

Sub SendBattleMessage(ByVal Index As Long, ByVal sMessage As String, Optional color As Long = White, Optional ByVal DontSendPVP As Long = 0)
On Error Resume Next
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteInteger SBattleMessage
    buffer.WriteString sMessage
    buffer.WriteLong color
    If TempPlayer(Index).isInPVP = True Then
    If DontSendPVP = 0 Then
    SendDataTo FindPlayer(Trim$(TempPlayer(Index).PVPEnemy)), buffer.ToArray()
    End If
    End If
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendDialog(ByVal Index As Long, ByVal dialog As String, Optional image As Long = 0, Optional isTrigger As Long = NO)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SDialogs
buffer.WriteString dialog
buffer.WriteLong image
buffer.WriteLong isTrigger
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendNPCScript(ByVal Index As Long, ByVal map As Long, ByVal x As Long, ByVal y As Long)
On Error Resume Next
Dim Name As String
Name = GetVar(App.Path & "\Data\NPCScripts\" & map & "I" & x & "I" & y & ".ini", "DATA", "Name")
Dim script As String
script = ReadText("Data\NPCScripts\" & map & "I" & x & "I" & y & ".txt")
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SNPCScript
buffer.WriteLong map
buffer.WriteLong x
buffer.WriteLong y
buffer.WriteString Name
buffer.WriteString script
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

'///////////////////////////////////////////////////////////////////////////
'//////Maximum packets reached now we will use one packet for all//////////
'/////////////////////////////////////////////////////////////////////////

Sub SendTo(ByVal Index As Long, ByVal packetType As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
'--------------------- Write packet type
buffer.WriteString packetType
'---------------------Write number of strings and longs
buffer.WriteLong 0
buffer.WriteLong 0
'--------------------Write Strings and longs (In this case its nothing but it needs to be written!)
buffer.WriteString ""
buffer.WriteLong 0
'---------------------
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub


Sub SendScript(ByVal Index As Long, ByVal script As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "NPCSCRIPT"
buffer.WriteLong 1
buffer.WriteLong 0
buffer.WriteString script
buffer.WriteLong 0
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendGoldNeeded(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "GOLDMSG"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString "a"
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendResetShop(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "RSHOP"
buffer.WriteLong 1
buffer.WriteLong 0
buffer.WriteString "A"
buffer.WriteLong 0
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendTradeStart(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "STARTTRADE"
buffer.WriteLong 1
buffer.WriteLong 0
buffer.WriteString Trim$(TempPlayer(Index).TradeName)
buffer.WriteLong 0
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendTradeStop(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "STOPTRADE"
buffer.WriteLong 1
buffer.WriteLong 0
buffer.WriteString "A"
buffer.WriteLong 0
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal poke As Long, ByVal item As Long, ByVal ItemVal As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "TRADEUPDATE"
buffer.WriteLong 1
buffer.WriteLong 11
buffer.WriteString "A"
If poke > 0 Then
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).PokemonNumber
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).level
Else
buffer.WriteLong 0
buffer.WriteLong 0
End If
If item > 0 Then
buffer.WriteLong GetPlayerInvItemNum(FindPlayer(Trim$(TempPlayer(Index).TradeName)), item)
If ItemVal = "" Then
buffer.WriteLong 1
Else
buffer.WriteLong Val(ItemVal)
End If
Else
buffer.WriteLong 0
buffer.WriteLong 0
End If
If poke > 0 Then
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).atk
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).def
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).spatk
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).spdef
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).spd
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).MaxHp
buffer.WriteLong player(FindPlayer(Trim$(TempPlayer(Index).TradeName))).PokemonInstance(poke).nature
Else
buffer.WriteLong 0
buffer.WriteLong 0
buffer.WriteLong 0
buffer.WriteLong 0
buffer.WriteLong 0
buffer.WriteLong 0
buffer.WriteLong 0
End If
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendTradeLocked(ByVal Index As Long, ByVal myacc As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "TRADELOCK"
buffer.WriteLong 1
buffer.WriteLong 2
buffer.WriteString "A"
If myacc = YES Then
buffer.WriteLong TempPlayer(Index).TradeLocked
Else
buffer.WriteLong TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeLocked
End If
buffer.WriteLong myacc
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendNews(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "NEWS"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString GetNews
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendLearnMove(ByVal Index As Long, ByVal poke As Long, ByVal move As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "LM"
buffer.WriteLong 1
buffer.WriteLong 2
buffer.WriteString ""
buffer.WriteLong poke
buffer.WriteLong move
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendEvolve(ByVal Index As Long, ByVal slot As Long, ByVal newPoke As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "EVOLVE"
buffer.WriteLong 1
buffer.WriteLong 2
buffer.WriteString ""
buffer.WriteLong slot
buffer.WriteLong newPoke
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendFlashlight(ByVal Index As Long, ByVal flashlight As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "FL"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString ""
buffer.WriteLong flashlight
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendTravel(ByVal Index As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "TRAVEL"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString ""
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendPCScan(ByVal Index As Long, ByVal admin As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "PCSCANREQUEST"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString admin
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendPCScanResultToAdmin(ByVal Index As Long, ByVal scanresult As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "PCSCANRESULT"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString scanresult
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendRadio(ByVal Index As Long, ByVal song As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "RADIO"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString song
buffer.WriteLong 1
SendDataToAll buffer.ToArray
Set buffer = Nothing
End Sub


Sub SendWhos(ByVal Index As Long, ByVal Num As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "WHOS"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString ""
buffer.WriteLong Num
SendDataToAll buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendCloseWhos(ByVal Index As Long, ByVal Num As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "CloseWHOS"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString ""
buffer.WriteLong 1
SendDataToAll buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendCrewData(ByVal Index As Long, ByVal crewname As String)
'On Error Resume Next
Dim i As Long
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SCrew
buffer.WriteString crewname
buffer.WriteString GetCrewPicture(crewname)
buffer.WriteString GetCrewLeaderName(crewname)
For i = 1 To 50
If Trim$(GetCrewPlayerName(crewname, i)) <> vbNullString Then
If FindPlayer(GetCrewPlayerName(crewname, i)) > 0 Then
If IsPlaying(FindPlayer(GetCrewPlayerName(crewname, i))) Then
buffer.WriteString GetCrewPlayerName(crewname, i) & " (Online)"
Else
buffer.WriteString GetCrewPlayerName(crewname, i) & " (Offline)"
End If
Else
buffer.WriteString GetCrewPlayerName(crewname, i) & " (Offline)"
End If
Else
buffer.WriteString "-"
End If
Next
buffer.WriteString GetCrewNews(crewname)
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendRemoveTP(ByVal Index As Long, ByVal slot As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger STPRemove
buffer.WriteLong slot
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
Sub SendPVPCommand(ByVal Index As Long, ByVal command As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SPVPCommand
buffer.WriteString command
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendJournal(ByVal Index As Long, ByVal playerIndex As Long)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SJournal
buffer.WriteString GetPlayerName(playerIndex)
buffer.WriteString GetPlayerJournal(playerIndex)
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendClanInvite(ByVal Index As Long, ByVal inviteIndex As Long, ByVal clanName As String)
On Error Resume Next
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "CLAN"
buffer.WriteLong 1
buffer.WriteLong 1
buffer.WriteString clanName
buffer.WriteLong inviteIndex
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub


Sub SendProfile(ByVal Index As Long)
On Error Resume Next
Dim dateNow As String
dateNow = Format(Date, "m/d/yyyy")
Dim Duration As Long
Dim dateStarted As String
Dim buffer As clsBuffer

Duration = GetPlayerMemberDuration(Index)
dateStarted = GetPlayerMemberDate(Index)

Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "PROFILE"
buffer.WriteLong 4
buffer.WriteLong 1
buffer.WriteString "Points@" & Trim$(GetRankedPoints(Index)) & "@"
buffer.WriteString "Membership@" & Trim$((Duration - DateDiff("d", dateStarted, dateNow))) & "@"
buffer.WriteString "Minutes@" & Trim$(GetPlayerPlaytimeMinutes(Index)) & "@"
buffer.WriteString "Hours@" & Trim$(GetPlayerPlaytimeHours(Index)) & "@"
'buffer.WriteLong GetRankedPoints(index) + 1
'buffer.WriteLong (Duration - DateDiff("d", dateStarted, dateNow)) + 1
'buffer.WriteLong GetPlayerPlaytimeMinutes(index) + 1
'buffer.WriteLong GetPlayerPlaytimeHours(index) + 1
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendEgg(ByVal Index As Long)
On Error Resume Next
Dim steps As Long
Dim expto As Long
Dim canHatch As String

Dim buffer As clsBuffer

steps = GetPlayerEggSteps(Index)
expto = GetPlayerEggEXP(Index)
If steps <= 0 And expto <= 0 Then
canHatch = "YES"
Else
canHatch = "NO"
End If

Set buffer = New clsBuffer
buffer.WriteInteger SSend
buffer.WriteString "EGG"
buffer.WriteLong 3
buffer.WriteLong 1
buffer.WriteString "Steps@" & Trim$(steps) & "@"
buffer.WriteString "Exp@" & Trim$(expto) & "@"
buffer.WriteString "Hatch@" & Trim$(canHatch) & "@"
buffer.WriteLong 1
SendDataTo Index, buffer.ToArray
Set buffer = Nothing
End Sub
