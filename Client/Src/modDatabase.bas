Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If

End Function

' gets a string from a text file
Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, file)
End Sub

Public Sub SaveOptions()
Dim FileName As String

FileName = App.Path & "\Data Files\config.ini"

Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
Call PutVar(FileName, "Options", "SavePass", str(Options.SavePass))
Call PutVar(FileName, "Options", "IP", Options.IP)
Call PutVar(FileName, "Options", "Port", str(Options.Port))
Call PutVar(FileName, "Options", "Music", Trim$(Options.music))
Call PutVar(FileName, "Options", "PlayMusic", Trim$(Options.PlayMusic))
Call PutVar(FileName, "Options", "RepeatMusic", Trim$(Options.repeatmusic))
Call PutVar(FileName, "Options", "CameraFollowPlayer", Trim$(Options.CameraFollowPlayer))
Call PutVar(FileName, "Options", "FormOpacity", Trim$(Options.FormTransparency))
Call PutVar(FileName, "Options", "PlayRadio", Trim$(Options.PlayRadio))
Call PutVar(FileName, "Options", "NearbyMaps", Trim$(Options.NearbyMaps))
End Sub

Public Sub LoadOptions()
Dim FileName As String

FileName = App.Path & "\Data Files\config.ini"

If Not FileExist(FileName, True) Then
    Options.Password = vbNullString
    Options.SavePass = 0
    Options.Username = vbNullString
    Options.IP = "127.0.0.1"
    Options.Port = 7001
    Options.music = vbNullString
    Options.PlayMusic = YES
    Options.repeatmusic = YES
    Options.CameraFollowPlayer = YES
    Options.FormTransparency = YES
    Options.NearbyMaps = YES
    SaveOptions
    Exit Sub
End If


Options.Username = GetVar(FileName, "Options", "Username")
Options.Password = GetVar(FileName, "Options", "Password")
Options.SavePass = Val(GetVar(FileName, "Options", "SavePass"))
Options.IP = GetVar(FileName, "Options", "IP")
Options.Port = Val(GetVar(FileName, "Options", "Port"))
Options.music = GetVar(FileName, "Options", "Music")
Options.repeatmusic = Val(GetVar(FileName, "Options", "RepeatMusic"))
Options.PlayMusic = Val(GetVar(FileName, "Options", "PlayMusic"))
Options.CameraFollowPlayer = Val(GetVar(FileName, "Options", "CameraFollowPlayer"))
Options.FormTransparency = Val(GetVar(FileName, "Options", "FormOpacity"))
Options.PlayRadio = Val(GetVar(FileName, "Options", "PlayRadio"))
Options.NearbyMaps = Val(GetVar(FileName, "Options", "NearbyMaps"))
End Sub

Public Sub AddLog(ByVal text As String)
    Dim FileName As String
    Dim f As Long

    'If DEBUG_MODE Then
    '    If Not frmDebug.Visible Then
    '        frmDebug.Visible = True
    '    End If

    '    FileName = App.Path & LOG_PATH & LOG_DEBUG

    '    If Not FileExist(LOG_DEBUG, True) Then
    '        f = FreeFile
    '        Open FileName For Output As #f
    '        Close #f
    '    End If

    '    f = FreeFile
    '    Open FileName For Append As #f
    '    Print #f, Time & ": " & Text
    '    Close #f
    'End If

End Sub

Public Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    'If FileExist(Filename, True) Then Kill Filename
    
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , map.Name
    Put #f, , map.Revision
    Put #f, , map.Moral
    Put #f, , map.tileset
    Put #f, , map.Up
    Put #f, , map.Down
    Put #f, , map.Left
    Put #f, , map.Right
    Put #f, , map.music
    Put #f, , map.BootMap
    Put #f, , map.BootX
    Put #f, , map.BootY
    Put #f, , map.MaxX
    Put #f, , map.MaxY

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            Put #f, , map.Tile(x, y)
        Next

        DoEvents
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , map.NPC(x)
    Next

    For x = 1 To MAX_MAP_POKEMONS
    Put #f, , map.Pokemon(x).PokemonNumber
    Put #f, , map.Pokemon(x).LevelFrom
    Put #f, , map.Pokemon(x).LevelTo
    Put #f, , map.Pokemon(x).Custom
    Put #f, , map.Pokemon(x).ATK
    Put #f, , map.Pokemon(x).DEF
    Put #f, , map.Pokemon(x).SPATK
    Put #f, , map.Pokemon(x).SPDEF
    Put #f, , map.Pokemon(x).SPD
    Put #f, , map.Pokemon(x).HP
    Put #f, , map.Pokemon(x).Chance
    Next
    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , map.Name
    Get #f, , map.Revision
    Get #f, , map.Moral
    Get #f, , map.tileset
    Get #f, , map.Up
    Get #f, , map.Down
    Get #f, , map.Left
    Get #f, , map.Right
    Get #f, , map.music
    Get #f, , map.BootMap
    Get #f, , map.BootX
    Get #f, , map.BootY
    Get #f, , map.MaxX
    Get #f, , map.MaxY
    ' have to set the tile()
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            Get #f, , map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , map.NPC(x)
    Next

 For x = 1 To MAX_MAP_POKEMONS
    Get #f, , map.Pokemon(x).PokemonNumber
    Get #f, , map.Pokemon(x).LevelFrom
    Get #f, , map.Pokemon(x).LevelTo
    Get #f, , map.Pokemon(x).Custom
    Get #f, , map.Pokemon(x).ATK
    Get #f, , map.Pokemon(x).DEF
    Get #f, , map.Pokemon(x).SPATK
    Get #f, , map.Pokemon(x).SPDEF
    Get #f, , map.Pokemon(x).SPD
    Get #f, , map.Pokemon(x).HP
    Get #f, , map.Pokemon(x).Chance
    Next

    Close #f
    ClearTempTile
    
    If map.Up > 0 Then
    Call LoadMapTo(map.Up, DIR_UP)
    End If
    
    If map.Down > 0 And map.Down <= MAX_MAPS Then
    Call LoadMapTo(map.Down, DIR_DOWN)
    End If
    
    If map.Left > 0 And map.Left <= MAX_MAPS Then
    Call LoadMapTo(map.Left, DIR_LEFT)
    End If
    
    If map.Right > 0 And map.Right <= MAX_MAPS Then
    Call LoadMapTo(map.Right, DIR_RIGHT)
    End If
End Sub

Public Sub CheckTilesets()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "\tilesets\" & i & GFX_EXT)
        NumTileSets = NumTileSets + 1
        i = i + 1
    Wend
    
    If NumTileSets = 0 Then Exit Sub

    frmEditor_Map.scrlTileSet.Max = NumTileSets
    
    ReDim DDS_Tileset(1 To NumTileSets)
    ReDim DDSD_Tileset(1 To NumTileSets)
    ReDim TilesetTimer(1 To NumTileSets)
End Sub

Public Sub CheckCharacters()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "characters\" & i & GFX_EXT)
        NumCharacters = NumCharacters + 1
        i = i + 1
    Wend
    
    If NumCharacters = 0 Then Exit Sub

    ReDim DDS_Character(1 To NumCharacters)
    ReDim DDSD_Character(1 To NumCharacters)
    ReDim CharacterTimer(1 To NumCharacters)
End Sub

Public Sub CheckPaperdolls()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "paperdolls\" & i & GFX_EXT)
        NumPaperdolls = NumPaperdolls + 1
        i = i + 1
    Wend
    
    If NumPaperdolls = 0 Then Exit Sub

    ReDim DDS_Paperdoll(1 To NumPaperdolls)
    ReDim DDSD_Paperdoll(1 To NumPaperdolls)
    ReDim PaperdollTimer(1 To NumPaperdolls)
End Sub

Public Sub CheckAnimations()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "animations\" & i & GFX_EXT)
        NumAnimations = NumAnimations + 1
        i = i + 1
    Wend
    
    If NumAnimations = 0 Then Exit Sub

    ReDim DDS_Animation(1 To NumAnimations)
    ReDim DDSD_Animation(1 To NumAnimations)
    ReDim AnimationTimer(1 To NumAnimations)
End Sub

Public Sub CheckItems()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "Items\" & i & GFX_EXT)
        NumItems = NumItems + 1
        i = i + 1
    Wend
    
    If NumItems = 0 Then Exit Sub

    ReDim DDS_Item(1 To NumItems)
    ReDim DDSD_Item(1 To NumItems)
    ReDim ItemTimer(1 To NumItems)
End Sub

Public Sub CheckResources()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "Resources\" & i & GFX_EXT)
        NumResources = NumResources + 1
        i = i + 1
    Wend
    
    If NumResources = 0 Then Exit Sub

    ReDim DDS_Resource(1 To NumResources)
    ReDim DDSD_Resource(1 To NumResources)
    ReDim ResourceTimer(1 To NumResources)
End Sub

Public Sub CheckSpellIcons()
    Dim i As Long
    i = 1

    While FileExist(GFX_PATH & "SpellIcons\" & i & GFX_EXT)
        NumSpellIcons = NumSpellIcons + 1
        i = i + 1
    Wend
    
    If NumSpellIcons = 0 Then Exit Sub

    ReDim DDS_SpellIcon(1 To NumSpellIcons)
    ReDim DDSD_SpellIcon(1 To NumSpellIcons)
    ReDim SpellIconTimer(1 To NumSpellIcons)
End Sub

Sub CheckOverWorld()
    
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
End Sub

Sub ClearItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
End Sub

Sub ClearAnimation(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

End Sub

Sub ClearNPC(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

Sub ClearShop(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

Sub ClearResource(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

End Sub

Sub ClearPokemon(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Pokemon(Index)), LenB(Pokemon(Index)))
    Pokemon(Index).Name = vbNullString
End Sub

Sub ClearPokemons()
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        Call ClearPokemon(i)
    Next

End Sub



Sub ClearMove(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(PokemonMove(Index)), LenB(PokemonMove(Index)))
    PokemonMove(Index).Name = vbNullString
End Sub

Sub ClearMoves()
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        Call ClearMove(i)
    Next

End Sub



Sub ClearMapItem(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
End Sub

Sub ClearMap()
    Call ZeroMemory(ByVal VarPtr(map), LenB(map))
    map.Name = vbNullString
    map.tileset = 1
    map.MaxX = MAX_MAPX
    map.MaxY = MAX_MAPY
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)
End Sub

Sub ClearNextMap(ByVal dir As Long)
Select Case dir
Case DIR_DOWN
Call ZeroMemory(ByVal VarPtr(DownMap), LenB(DownMap))
    DownMap.Name = vbNullString
    DownMap.tileset = 1
    DownMap.MaxX = MAX_MAPX
    DownMap.MaxY = MAX_MAPY
    ReDim DownMap.Tile(0 To DownMap.MaxX, 0 To DownMap.MaxY)
Case DIR_UP
Call ZeroMemory(ByVal VarPtr(UpMap), LenB(UpMap))
    UpMap.Name = vbNullString
   UpMap.tileset = 1
    UpMap.MaxX = MAX_MAPX
    UpMap.MaxY = MAX_MAPY
    ReDim UpMap.Tile(0 To UpMap.MaxX, 0 To UpMap.MaxY)
Case DIR_LEFT
Call ZeroMemory(ByVal VarPtr(LeftMap), LenB(LeftMap))
    LeftMap.Name = vbNullString
   LeftMap.tileset = 1
    LeftMap.MaxX = MAX_MAPX
    LeftMap.MaxY = MAX_MAPY
    ReDim LeftMap.Tile(0 To LeftMap.MaxX, 0 To LeftMap.MaxY)
Case DIR_RIGHT
Call ZeroMemory(ByVal VarPtr(RightMap), LenB(RightMap))
    RightMap.Name = vbNullString
   RightMap.tileset = 1
    RightMap.MaxX = MAX_MAPX
    RightMap.MaxY = MAX_MAPY
    ReDim RightMap.Tile(0 To RightMap.MaxX, 0 To RightMap.MaxY)

End Select
    
End Sub


Sub ClearMapItems()
    Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
On Error Resume Next
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long

    If Index > MAX_PLAYERS Then Exit Function

    Select Case Vital
        Case HP
            GetPlayerMaxVital = Player(Index).MaxHp
        Case MP
            GetPlayerMaxVital = Player(Index).MaxMP
        Case SP
            GetPlayerMaxVital = Player(Index).MaxSP
    End Select

End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal dir As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).dir = dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invslot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invslot As Long, ByVal itemnum As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).num = itemnum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long) As Long

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invslot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invslot As Long, ByVal itemvalue As Long)

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invslot).Value = itemvalue
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invnum As Long, ByVal EquipmentSlot As Equipment)

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invnum
End Sub
Sub LoadNature()
    Dim FileName As String
    Dim i As Integer
    FileName = App.Path & "\Data Files\database\Natures.ini"
    For i = 1 To MAX_NATURES
        nature(i).Name = GetVar(FileName, "NATURE" & i, "Name")
        nature(i).AddHP = Val(GetVar(FileName, "NATURE" & i, "HP"))
        nature(i).AddAtk = Val(GetVar(FileName, "NATURE" & i, "ATK"))
        nature(i).AddDef = Val(GetVar(FileName, "NATURE" & i, "DEF"))
        nature(i).AddSpAtk = Val(GetVar(FileName, "NATURE" & i, "SPATK"))
        nature(i).AddSpDef = Val(GetVar(FileName, "NATURE" & i, "SPDEF"))
        nature(i).AddSpd = Val(GetVar(FileName, "NATURE" & i, "SPEED"))
    Next
End Sub


Public Sub LoadMapTo(ByVal MapNum As Long, ByVal dir As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearNextMap dir
    f = FreeFile
    Select Case dir
    Case DIR_UP
    Open FileName For Binary As #f
    Get #f, , UpMap.Name
    Get #f, , UpMap.Revision
    Get #f, , UpMap.Moral
    Get #f, , UpMap.tileset
    Get #f, , UpMap.Up
    Get #f, , UpMap.Down
    Get #f, , UpMap.Left
    Get #f, , UpMap.Right
    Get #f, , UpMap.music
    Get #f, , UpMap.BootMap
    Get #f, , UpMap.BootX
    Get #f, , UpMap.BootY
    Get #f, , UpMap.MaxX
    Get #f, , UpMap.MaxY
    ' have to set the tile()
    ReDim UpMap.Tile(0 To UpMap.MaxX, 0 To UpMap.MaxY)

    For x = 0 To UpMap.MaxX
        For y = 0 To UpMap.MaxY
            Get #f, , UpMap.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , UpMap.NPC(x)
    Next

 For x = 1 To MAX_MAP_POKEMONS
    Get #f, , UpMap.Pokemon(x).PokemonNumber
    Get #f, , UpMap.Pokemon(x).LevelFrom
    Get #f, , UpMap.Pokemon(x).LevelTo
    Get #f, , UpMap.Pokemon(x).Custom
    Get #f, , UpMap.Pokemon(x).ATK
    Get #f, , UpMap.Pokemon(x).DEF
    Get #f, , UpMap.Pokemon(x).SPATK
    Get #f, , UpMap.Pokemon(x).SPDEF
    Get #f, , UpMap.Pokemon(x).SPD
    Get #f, , UpMap.Pokemon(x).HP
    Get #f, , UpMap.Pokemon(x).Chance
    Next

    Close #f
    
    Case DIR_DOWN
    
     Open FileName For Binary As #f
    Get #f, , DownMap.Name
    Get #f, , DownMap.Revision
    Get #f, , DownMap.Moral
    Get #f, , DownMap.tileset
    Get #f, , DownMap.Up
    Get #f, , DownMap.Down
    Get #f, , DownMap.Left
    Get #f, , DownMap.Right
    Get #f, , DownMap.music
    Get #f, , DownMap.BootMap
    Get #f, , DownMap.BootX
    Get #f, , DownMap.BootY
    Get #f, , DownMap.MaxX
    Get #f, , DownMap.MaxY
    ' have to set the tile()
    ReDim DownMap.Tile(0 To DownMap.MaxX, 0 To DownMap.MaxY)

    For x = 0 To DownMap.MaxX
        For y = 0 To DownMap.MaxY
            Get #f, , DownMap.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , DownMap.NPC(x)
    Next

 For x = 1 To MAX_MAP_POKEMONS
    Get #f, , DownMap.Pokemon(x).PokemonNumber
    Get #f, , DownMap.Pokemon(x).LevelFrom
    Get #f, , DownMap.Pokemon(x).LevelTo
    Get #f, , DownMap.Pokemon(x).Custom
    Get #f, , DownMap.Pokemon(x).ATK
    Get #f, , DownMap.Pokemon(x).DEF
    Get #f, , DownMap.Pokemon(x).SPATK
    Get #f, , DownMap.Pokemon(x).SPDEF
    Get #f, , DownMap.Pokemon(x).SPD
    Get #f, , DownMap.Pokemon(x).HP
    Get #f, , DownMap.Pokemon(x).Chance
    Next

    Close #f
    
    Case DIR_LEFT
    
     Open FileName For Binary As #f
    Get #f, , LeftMap.Name
    Get #f, , LeftMap.Revision
    Get #f, , LeftMap.Moral
    Get #f, , LeftMap.tileset
    Get #f, , LeftMap.Up
    Get #f, , LeftMap.Down
    Get #f, , LeftMap.Left
    Get #f, , LeftMap.Right
    Get #f, , LeftMap.music
    Get #f, , LeftMap.BootMap
    Get #f, , LeftMap.BootX
    Get #f, , LeftMap.BootY
    Get #f, , LeftMap.MaxX
    Get #f, , LeftMap.MaxY
    ' have to set the tile()
    ReDim LeftMap.Tile(0 To LeftMap.MaxX, 0 To LeftMap.MaxY)

    For x = 0 To LeftMap.MaxX
        For y = 0 To LeftMap.MaxY
            Get #f, , LeftMap.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , LeftMap.NPC(x)
    Next

 For x = 1 To MAX_MAP_POKEMONS
    Get #f, , LeftMap.Pokemon(x).PokemonNumber
    Get #f, , LeftMap.Pokemon(x).LevelFrom
    Get #f, , LeftMap.Pokemon(x).LevelTo
    Get #f, , LeftMap.Pokemon(x).Custom
    Get #f, , LeftMap.Pokemon(x).ATK
    Get #f, , LeftMap.Pokemon(x).DEF
    Get #f, , LeftMap.Pokemon(x).SPATK
    Get #f, , LeftMap.Pokemon(x).SPDEF
    Get #f, , LeftMap.Pokemon(x).SPD
    Get #f, , LeftMap.Pokemon(x).HP
    Get #f, , LeftMap.Pokemon(x).Chance
    Next

    Close #f
    
    
    Case DIR_RIGHT
    
    Open FileName For Binary As #f
    Get #f, , RightMap.Name
    Get #f, , RightMap.Revision
    Get #f, , RightMap.Moral
    Get #f, , RightMap.tileset
    Get #f, , RightMap.Up
    Get #f, , RightMap.Down
    Get #f, , RightMap.Left
    Get #f, , RightMap.Right
    Get #f, , RightMap.music
    Get #f, , RightMap.BootMap
    Get #f, , RightMap.BootX
    Get #f, , RightMap.BootY
    Get #f, , RightMap.MaxX
    Get #f, , RightMap.MaxY
    ' have to set the tile()
    ReDim RightMap.Tile(0 To RightMap.MaxX, 0 To RightMap.MaxY)

    For x = 0 To RightMap.MaxX
        For y = 0 To RightMap.MaxY
            Get #f, , RightMap.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , RightMap.NPC(x)
    Next

 For x = 1 To MAX_MAP_POKEMONS
    Get #f, , RightMap.Pokemon(x).PokemonNumber
    Get #f, , RightMap.Pokemon(x).LevelFrom
    Get #f, , RightMap.Pokemon(x).LevelTo
    Get #f, , RightMap.Pokemon(x).Custom
    Get #f, , RightMap.Pokemon(x).ATK
    Get #f, , RightMap.Pokemon(x).DEF
    Get #f, , RightMap.Pokemon(x).SPATK
    Get #f, , RightMap.Pokemon(x).SPDEF
    Get #f, , RightMap.Pokemon(x).SPD
    Get #f, , RightMap.Pokemon(x).HP
    Get #f, , RightMap.Pokemon(x).Chance
    Next

    Close #f
    
    
    End Select

End Sub

Public Sub LoadBattleGDI()

If PokemonInstance(BattlePokemon).isShiny = YES Then
Set PokeImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\Back\" & PokemonInstance(BattlePokemon).PokemonNumber & ".gif")
Else
Set PokeImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Back\" & PokemonInstance(BattlePokemon).PokemonNumber & ".gif")
End If
If enemyPokemon.isShiny = YES Then
Set EnemyPokeImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & enemyPokemon.PokemonNumber & ".gif")
Else
Set EnemyPokeImg = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & enemyPokemon.PokemonNumber & ".gif")
End If
PokeImgLoaded = True
EnemyPokeImgLoaded = True
End Sub
