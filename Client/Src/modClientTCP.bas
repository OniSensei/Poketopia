Attribute VB_Name = "modClientTCP"
Option Explicit
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer

Sub TcpInit()
    Set PlayerBuffer = New clsBuffer
    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)

    ' check if IP is valid
    If IsIP(Options.IP) Then
        frmMainGame.Socket.RemoteHost = Options.IP
        frmMainGame.Socket.RemotePort = Options.Port
    Else
        MsgBox Options.IP & " does not appear as a valid IP address!"
        DestroyGame
    End If
   End Sub

Sub DestroyTCP()
    frmMainGame.Socket.Close
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long
    Dim Data() As Byte
    frmMainGame.Socket.GetData Buffer, vbUnicode, DataLength
    PlayerBuffer.WriteBytes Buffer()

    If PlayerBuffer.Length >= 4 Then
        pLength = PlayerBuffer.ReadLong(False)
    End If

    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4

        'make sure we have the right plength and pbuffer
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            Data() = PlayerBuffer.ReadBytes(pLength + 1)
            HandleData Data()
        End If

        pLength = 0

        If PlayerBuffer.Length >= 4 Then
            pLength = PlayerBuffer.ReadLong(False)
        End If

    Loop

    ' Check if the playbuffer is empty
    If PlayerBuffer.Length <= 1 Then PlayerBuffer.Flush
End Sub

Public Function ConnectToServer(ByVal i As Long) As Boolean
    Dim Wait As Long

    ' Check to see if we are already connected, if so just exit
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    If i = 2 Then Exit Function
    Wait = GetTickCount

    With frmMainGame.Socket
        .Close
        .Connect
    End With

    Call SetStatus("Connecting to server...(" & i & ")")

    ' Wait until connected or a few seconds have passed and report the server being down
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3500)
        DoEvents
        Sleep 20
    Loop

    ' return value
    If IsConnected Then
        ConnectToServer = True
    End If

    If Not ConnectToServer Then
        Call ConnectToServer(i + 1)
    End If

End Function

Private Function IsIP(ByVal IPAddress As String) As Boolean
    Dim S() As String
    Dim i As Long

    ' Check if connecting to localhost or URL
    If IPAddress = "localhost" Or InStr(1, IPAddress, "http://", vbTextCompare) = 1 Then
        IsIP = True
        Exit Function
    End If

    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    'Split up the string by the periods
    S = Split(IPAddress, ".")

    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(S) <> 3 Then Exit Function

    'Check that the values are numeric and in a valid range
    For i = 0 To 3

        If Val(S(i)) < 0 Then Exit Function
        If Val(S(i)) > 255 Then Exit Function
    Next

    'Looks like we were passed a valid IP!
    IsIP = True
End Function

Function IsConnected() As Boolean

    If frmMainGame.Socket.State = sckConnected Then
        IsConnected = True
    End If

End Function

Function IsPlaying(ByVal Index As Long) As Boolean

    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

End Function

Sub SendData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    If IsConnected Then
        Set Buffer = New clsBuffer
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data
    
        frmMainGame.Socket.SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNewAccount
    Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelAccount
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    Buffer.WriteString Password
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    Buffer.WriteString Password
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long, ByVal Starter As Long, ByVal hairC As Long, ByVal hairI As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    Buffer.WriteLong Sex
    Buffer.WriteLong ClassNum
    Buffer.WriteLong Sprite
    Buffer.WriteLong Starter
    Buffer.WriteLong hairC
    Buffer.WriteLong hairI
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseChar
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong CharSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPCScan(ByVal playerName As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPCScan
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString playerName
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPCScanResult()
Dim i As Long
Dim a As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPCScan
     Buffer.WriteLong TCP_CODE
    Dim process As Object
For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
a = a + 1
Next
Buffer.WriteLong a
For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
    Buffer.WriteString process.Caption
Next
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub



Public Sub SayMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BroadcastMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBroadcastMsg
    Buffer.WriteLong TCP_CODE
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub EmoteMsg(ByVal text As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CEmoteMsg
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMood(ByVal moodState As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CSetMood
 Buffer.WriteLong TCP_CODE
Buffer.WriteLong moodState
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal text As String, ByVal MsgTo As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayMsg
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString MsgTo
    Buffer.WriteString text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerMove()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong GetPlayerDir(MyIndex)
    Buffer.WriteLong Player(MyIndex).Moving
    Buffer.WriteLong Player(MyIndex).x
    Buffer.WriteLong Player(MyIndex).y
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerDir()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPlayerRequestNewMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMap()
    Dim packet As String
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    CanMoveNow = False

    With map
        Buffer.WriteLong CMapData
         Buffer.WriteLong TCP_CODE
        Buffer.WriteString Trim$(.Name)
        Buffer.WriteLong .Moral
        Buffer.WriteLong .tileset
        Buffer.WriteLong .Up
        Buffer.WriteLong .Down
        Buffer.WriteLong .Left
        Buffer.WriteLong .Right
        Buffer.WriteLong .music
        Buffer.WriteLong .BootMap
        Buffer.WriteLong .BootX
        Buffer.WriteLong .BootY
        Buffer.WriteLong .MaxX
        Buffer.WriteLong .MaxY
    End With

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY

            With map.Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteByte .Layer(i).x
                    Buffer.WriteByte .Layer(i).y
                    Buffer.WriteByte .Layer(i).tileset
                Next
                Buffer.WriteLong .Type
                Buffer.WriteLong .data1
                Buffer.WriteLong .data2
                Buffer.WriteLong .data3
            End With

        Next
    Next

    With map

        For x = 1 To MAX_MAP_NPCS
            Buffer.WriteLong .NPC(x)
        Next

    End With

    With map
       For x = 1 To MAX_MAP_POKEMONS
       Buffer.WriteLong .Pokemon(x).PokemonNumber
       Buffer.WriteLong .Pokemon(x).LevelFrom
       Buffer.WriteLong .Pokemon(x).LevelTo
       Buffer.WriteLong .Pokemon(x).Custom
       Buffer.WriteLong .Pokemon(x).ATK
       Buffer.WriteLong .Pokemon(x).DEF
       Buffer.WriteLong .Pokemon(x).SPATK
       Buffer.WriteLong .Pokemon(x).SPDEF
       Buffer.WriteLong .Pokemon(x).SPD
       Buffer.WriteLong .Pokemon(x).HP
       Buffer.WriteLong .Pokemon(x).Chance
       Next
    End With
    
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub


Public Sub WarpAdmin()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpAdmin
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong CurX
    Buffer.WriteLong CurY
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpMeTo
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpToMe
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWarpTo
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong MapNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetAccess
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    Buffer.WriteLong Access
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetSprite
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong SpriteNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendKick(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CKickPlayer
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBan(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanPlayer
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanList()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanList
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditItem()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveItem(ByVal itemnum As Long)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    Buffer.WriteLong CSaveItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong itemnum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditAnimation()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditAnimation
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    Buffer.WriteLong CSaveAnimation
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong Animationnum
    Buffer.WriteBytes AnimationData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditNpc()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNpc
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveNpc(ByVal npcnum As Long)
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    NpcSize = LenB(NPC(npcnum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcnum)), NpcSize
    Buffer.WriteLong CSaveNpc
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong npcnum
    Buffer.WriteBytes NpcData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditResource()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditResource
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    Set Buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    Buffer.WriteLong CSaveResource
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditPokemon()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPokemon
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Public Sub SendRequestEditMove()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMove
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSavePokemon(ByVal pokemonnum As Long)
    Dim Buffer As clsBuffer
    Dim Pokemonize As Long
    Dim PokemonData() As Byte
    Set Buffer = New clsBuffer
    Pokemonize = LenB(Pokemon(pokemonnum))
    ReDim PokemonData(Pokemonize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(pokemonnum)), Pokemonize
    Buffer.WriteLong CSavePokemon
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong pokemonnum
    Buffer.WriteBytes PokemonData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveMove(ByVal movenum As Long)
    Dim Buffer As clsBuffer
    Dim MoveSize As Long
    Dim MoveData() As Byte
    Set Buffer = New clsBuffer
    MoveSize = LenB(PokemonMove(movenum))
    ReDim MoveData(MoveSize - 1)
    CopyMemory MoveData(0), ByVal VarPtr(PokemonMove(movenum)), MoveSize
    Buffer.WriteLong CSaveMove
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong movenum
    Buffer.WriteBytes MoveData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMapRespawn()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapRespawn
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendUseItem(ByVal invnum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong invnum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendDropItem(ByVal invnum As Long, ByVal Amount As Long)
    Dim Buffer As clsBuffer
    
    ' do basic checks
    If invnum < 1 Or invnum > MAX_INV Then Exit Sub
    If PlayerInv(invnum).num < 1 Or PlayerInv(invnum).num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invnum)).Type = ITEM_TYPE_CURRENCY Then
        If Amount < 1 Or Amount > PlayerInv(invnum).Value Then Exit Sub
    End If
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMapDropItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong invnum
    Buffer.WriteLong Amount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SetMapMusic(ByVal music As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CSetMapMusic
 Buffer.WriteLong TCP_CODE
Buffer.WriteString music
SendData Buffer.ToArray()
Set Buffer = Nothing
End Sub

Public Sub SendWhosOnline()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWhosOnline
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSetMotd
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString MOTD
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditShop()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    Buffer.WriteLong CSaveShop
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong shopnum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditSpell()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditSpell
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    Buffer.WriteLong CSaveSpell
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong spellnum
    Buffer.WriteBytes SpellData
    SendData Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Public Sub SendRequestEditMap()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendPartyRequest(ByVal Name As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CParty
     Buffer.WriteLong TCP_CODE
    Buffer.WriteString Name
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendJoinParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CJoinParty
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendLeaveParty()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLeaveParty
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBanDestroy()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBanDestroy
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Integer, ByVal NewSlot As Integer)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwapInvSlots
     Buffer.WriteLong TCP_CODE
    Buffer.WriteInteger OldSlot
    Buffer.WriteInteger NewSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub GetPing()
    Dim Buffer As clsBuffer
    PingStart = GetTickCount
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCheckPing
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUnequip(ByVal EqNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUnequip
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong EqNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestPlayerData()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPlayerData
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestItems()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestAnimations()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestAnimations
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestNPCS()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCS
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestResources()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestResources
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestPokemon()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPokemon
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub



Sub SendRequestMove()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestMove
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDepositPC()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CDepositPC
 Buffer.WriteLong TCP_CODE
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendWithdrawPC()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CWithdrawPC
 Buffer.WriteLong TCP_CODE
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendDepositPokemon(ByVal pokemonslot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositPokemon
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong pokemonslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSetAsLeader(ByVal pokeslot As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CSetAsLeader
 Buffer.WriteLong TCP_CODE
Buffer.WriteLong pokeslot
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub


Sub SendAddTP(ByVal Stat As Long, ByVal pokeslot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddTP
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong Stat
    Buffer.WriteLong pokeslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRosterRequest()
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CRosterRequest
 Buffer.WriteLong TCP_CODE
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendWithdrawPokemon(ByVal storageslot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawPokemon
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong storageslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRemoveStoragePokemon(ByVal storageslot As Long)
Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRemoveStoragePokemon
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong storageslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestSpells()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestSpells
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendRequestShops()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSpawnItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong tmpItem
    Buffer.WriteLong tmpAmount
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CTrainStat
     Buffer.WriteLong TCP_CODE
    Buffer.WriteByte StatNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendRequestLevelUp()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestLevelUp
     Buffer.WriteLong TCP_CODE
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub BuyItem(ByVal shopslot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong shopslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SellItem(ByVal invslot As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
     Buffer.WriteLong TCP_CODE
    Buffer.WriteLong invslot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendBattleCommand(ByVal Command As Byte, ByVal pokeslot As Long, ByVal move As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBattleCommand
     Buffer.WriteLong TCP_CODE
    Buffer.WriteByte Command
    Buffer.WriteLong pokeslot
    Buffer.WriteLong move
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendEditMapNpc(ByVal npcnum As Long, ByVal script As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CMapNPC
 Buffer.WriteLong TCP_CODE
Buffer.WriteLong npcnum
Buffer.WriteString script
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendRequest(ByVal data1 As Long, ByVal data2 As Long, ByVal data3 As String, ByVal reqType As String, Optional ByVal Data4 As String = "")
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CRequests
 Buffer.WriteLong TCP_CODE
Buffer.WriteLong data1
Buffer.WriteLong data2
Buffer.WriteString data3
Buffer.WriteString Data4
Buffer.WriteString reqType
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub

Sub SendLearnMove(ByVal pokeslot As Long, ByVal moveSlot As Long, ByVal move As Long)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CLearnMove
 Buffer.WriteLong TCP_CODE
Buffer.WriteLong pokeslot
Buffer.WriteLong moveSlot
Buffer.WriteLong move
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub
Sub SendDonate(ByVal optionUsed As String, ByVal ID As String, ByVal realName As String, ByVal email As String)
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteLong CDonate
 Buffer.WriteLong TCP_CODE
Buffer.WriteString optionUsed
Buffer.WriteString ID
Buffer.WriteString realName
Buffer.WriteString email
SendData Buffer.ToArray
Set Buffer = Nothing
End Sub
