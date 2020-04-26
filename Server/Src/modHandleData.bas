Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
On Error Resume Next
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
On Error Resume Next
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CParty) = GetAddress(AddressOf HandleParty)
    HandleDataSub(CJoinParty) = GetAddress(AddressOf HandleJoinParty)
    HandleDataSub(CLeaveParty) = GetAddress(AddressOf HandleLeaveParty)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CRequestEditPokemon) = GetAddress(AddressOf HandleRequestEditPokemon)
    HandleDataSub(CSavePokemon) = GetAddress(AddressOf HandleSavePokemon)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CRequestPokemon) = GetAddress(AddressOf HandleRequestPokemon)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CTrainStat) = GetAddress(AddressOf HandleTrainStat)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CBattleCommand) = GetAddress(AddressOf HandleBattleCommand)
    HandleDataSub(CSaveMove) = GetAddress(AddressOf HandleSaveMove)
    HandleDataSub(CRequestMove) = GetAddress(AddressOf HandleRequestMove)
    HandleDataSub(CRequestEditMove) = GetAddress(AddressOf HandleRequestEditMove)
    HandleDataSub(CDepositPokemon) = GetAddress(AddressOf HandleDepositPokemon)
    HandleDataSub(CWithdrawPokemon) = GetAddress(AddressOf HandleWithdrawPokemon)
    HandleDataSub(CDepositPC) = GetAddress(AddressOf HandleDepositPC)
    HandleDataSub(CWithdrawPC) = GetAddress(AddressOf HandleWithdrawPC)
    HandleDataSub(CRemoveStoragePokemon) = GetAddress(AddressOf HandleRemoveStoragePokemon)
    HandleDataSub(CAddTP) = GetAddress(AddressOf HandleAddTP)
    HandleDataSub(CRosterRequest) = GetAddress(AddressOf HandleRosterRequest)
    HandleDataSub(CSetAsLeader) = GetAddress(AddressOf HandleSetAsLeader)
    HandleDataSub(CWarpAdmin) = GetAddress(AddressOf HandleWarpAdmin)
    HandleDataSub(CPCScan) = GetAddress(AddressOf HandlePCScan)
    HandleDataSub(CPCScanResult) = GetAddress(AddressOf HandlePCScanResult)
    HandleDataSub(CSetMood) = GetAddress(AddressOf HandleSetMood)
    HandleDataSub(CSetMapMusic) = GetAddress(AddressOf HandleSetMapMusic)
    HandleDataSub(CMapNPC) = GetAddress(AddressOf HandleMapNPC)
    HandleDataSub(CRequests) = GetAddress(AddressOf HandleRequests)
    HandleDataSub(CLearnMove) = GetAddress(AddressOf HandleLearnMove)
    HandleDataSub(CDonate) = GetAddress(AddressOf HandleDonate)
End Sub

' Will handle the packet data
Sub HandleData(ByVal Index As Long, ByRef Data() As Byte)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim MsgType As Long
    Dim tcpValue As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong
    tcpValue = buffer.ReadLong
    If MsgType < 0 Then
        Exit Sub
    End If

    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    If tcpValue = TCP_CODE Then

    Else
   
    
    ServerBanIndex Index, "You are using non familiar 3rd party client.Your account has been blocked.Contact staff for more information."
    End If

    CallWindowProc HandleDataSub(MsgType), Index, buffer.ReadBytes(buffer.Length - 4 + 1), 0, 0
    Set buffer = Nothing
End Sub

Private Sub HandleNewAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(Index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(player(Index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar Index
                Else
                    ' send new char shit
                    If Not IsPlaying(Index) Then
                        Call SendNewCharClasses(Index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
            
            Set buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(Index, Name)

            If LenB(Trim$(player(Index).Name)) > 0 Then
                Call DeleteName(player(Index).Name)
            End If

            Call ClearPlayer(Index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        If Not IsLoggedIn(Index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Check versions
            If buffer.ReadLong < CLIENT_MAJOR Or buffer.ReadLong < CLIENT_MINOR Or buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(Index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(Index, "Multiple account logins is not authorized.")
                Exit Sub
            End If
            
            'If GetVar(App.Path & "\Data\Golds.ini", "Players", Name) = "YES" Then
            'Else
            'Call AlertMsg(index, "Login is enabled for gold players only during the BETA testing.Please check our facebook page for more and find out when next all players testing will be!")
            'Call SendGoldNeeded(index)
            'Exit Sub
            'End If
            
           
            

            ' Load the player
            Call LoadPlayer(Index, Name)
            
            If GetVar(App.Path & "\Data\banlist.ini", "DATA", Trim$(player(Index).Name)) = "YES" Then
            Call AlertMsg(Index, "You have been banned!")
            End If
            
             If AdminOnly = True Then
            If player(Index).Access < 1 Then
            Call AlertMsg(Index, "Server is up for admins only.Please stay tuned.")
            
            Exit Sub
            End If
            End If
            
            
            
            ' Check if character data has been created
            If LenB(Trim$(player(Index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar Index
            If GetVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Logged") <> "YES" Then
            SendIntro Index
            Call PutVar(App.Path & "\Data\alive\" & Trim$(player(Index).Name) & ".ini", "Other", "Logged", "YES")
            End If
            Else
                ' send new char shit
                If Not IsPlaying(Index) Then
                    Call SendNewCharClasses(Index)
                End If
            End If
            
           
            
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".")
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim Starter As Long
    Dim hairC As Long
    Dim hairI As Long
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(Index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Name = buffer.ReadString
        Sex = buffer.ReadLong
        Class = buffer.ReadLong
        Sprite = buffer.ReadLong
        Starter = buffer.ReadLong
        hairC = buffer.ReadLong
        hairI = buffer.ReadLong
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(Index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(Index) Then
            Call AlertMsg(Index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(Index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(Index, Name, Sex, Class, Sprite, Starter, hairC, hairI)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
        ' log them in!!
        If GoldNeeded Then
        If GetVar(App.Path & "\Data\Golds.ini", "Players", Name) = "YES" Then
        HandleUseChar Index
        SendIntro Index
        Call PutVar(App.Path & "\Data\alive\" & Name & ".ini", "Other", "Logged", "YES")
        Else
        AlertMsg Index, "Only choosen members are allowed to login! Please contact our staff!"
        
        End If
        Else
        HandleUseChar Index
        SendIntro Index
        Call PutVar(App.Path & "\Data\alive\" & Name & ".ini", "Other", "Logged", "YES")
        End If
        
        
        
        Set buffer = Nothing
    End If

End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(Index), Index, msg, QBColor(White))
    Call SendActionMsg(GetPlayerMap(Index), Trim$(msg), White, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
    Set buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim msg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Right$(msg, Len(msg) - 1), EmoteColor)
    
    Set buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim msg As String
    Dim s As String
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next
    If isPlayerMuted(Index) = False Then
    s = "[Global]" & GetPlayerName(Index) & ": " & msg
    Call SayMsg_Global(Index, msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    End If
    
    Set buffer = Nothing
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgTo = FindPlayer(buffer.ReadString)
    msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(msg)

        If AscW(Mid$(msg, i, 1)) < 32 Or AscW(Mid$(msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> Index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & msg & "'", TellColor)
            Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & msg & "'", TellColor)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(GetPlayerName(Index), "Cannot message yourself.", BrightRed)
    End If
    
    Set buffer = Nothing

End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Dir As Long
    Dim Movement As Long
    Dim buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    Movement = buffer.ReadLong 'CLng(Parse(2))
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(Index).SpellBuffer > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' if stunned, stop them moving
    If TempPlayer(Index).StunDuration > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(Index).InShop > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If
    
    ' prevent player from moiving if in battle
    If TempPlayer(Index).BattleType > 0 Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(Index) <> tmpX Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    If GetPlayerY(Index) <> tmpY Then
        SendPlayerXY (Index)
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, Movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(Index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    Set buffer = New clsBuffer
    buffer.WriteInteger SPlayerDir
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerDir(Index)
    SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim InvNum As Long
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim buffer As clsBuffer
    Dim TempItem As Long ' hold a currently equiped item data :]
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvNum = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
        n = item(GetPlayerInvItemNum(Index, InvNum)).Data2

        ' Find out what kind of item it is
        Select Case item(GetPlayerInvItemNum(Index, InvNum)).Type
            Case ITEM_TYPE_ARMOR
            
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Armor) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Armor)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Armor
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 0 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
            Case ITEM_TYPE_WEAPON
            
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Weapon) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Weapon)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Weapon
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 0 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
            Case ITEM_TYPE_HELMET
            
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Helmet) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Helmet)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Helmet
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 0 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
            Case ITEM_TYPE_SHIELD
            
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Shield) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Shield)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Shield
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 1 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
            Case ITEM_TYPE_MASK
             For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Equipment.Mask) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Equipment.Mask)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Equipment.Mask
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 1 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
                
                Case ITEM_TYPE_OUTFIT
             For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to equip this item.", BrightRed
                        Exit Sub
                    End If
                Next

                If GetPlayerEquipment(Index, Equipment.Outfit) > 0 Then
                    TempItem = GetPlayerEquipment(Index, Equipment.Outfit)
                End If

                SetPlayerEquipment Index, GetPlayerInvItemNum(Index, InvNum), Equipment.Outfit
                PlayerMsg Index, "You equip " & CheckGrammar(item(GetPlayerInvItemNum(Index, InvNum)).Name), BrightGreen
                TakeItem Index, GetPlayerInvItemNum(Index, InvNum), 1

                If TempItem > 0 Then
                    GiveItem Index, TempItem, 1 ' give back the stored item
                    TempItem = 0
                End If

                Call SendWornEquipment(Index)
                Call SendMapEquipment(Index)
            
            
            Case ITEM_TYPE_POTIONADDHP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                SendActionMsg GetPlayerMap(Index), "+" & item(player(Index).Inv(InvNum).Num).Data1, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.hp, GetPlayerVital(Index, Vitals.hp) + item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.hp)
            Case ITEM_TYPE_POTIONADDMP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                SendActionMsg GetPlayerMap(Index), "+" & item(player(Index).Inv(InvNum).Num).Data1, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.mp, GetPlayerVital(Index, Vitals.mp) + item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.mp)
            Case ITEM_TYPE_POTIONADDSP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) + item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.SP)
            Case ITEM_TYPE_POTIONSUBHP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                SendActionMsg GetPlayerMap(Index), "-" & item(player(Index).Inv(InvNum).Num).Data1, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.hp, GetPlayerVital(Index, Vitals.hp) - item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.hp)
            Case ITEM_TYPE_POTIONSUBMP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                SendActionMsg GetPlayerMap(Index), "-" & item(player(Index).Inv(InvNum).Num).Data1, Blue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.mp, GetPlayerVital(Index, Vitals.mp) - item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.mp)
            Case ITEM_TYPE_POTIONSUBSP
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                Call SetPlayerVital(Index, Vitals.SP, GetPlayerVital(Index, Vitals.SP) - item(player(Index).Inv(InvNum).Num).Data1)
                Call TakeItem(Index, player(Index).Inv(InvNum).Num, 0)
                Call SendVital(Index, Vitals.SP)
            Case ITEM_TYPE_KEY
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next

                Select Case GetPlayerDir(Index)
                    Case DIR_UP

                        If GetPlayerY(Index) > 0 Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) - 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_DOWN

                        If GetPlayerY(Index) < map(GetPlayerMap(Index)).MaxY Then
                            x = GetPlayerX(Index)
                            y = GetPlayerY(Index) + 1
                        Else
                            Exit Sub
                        End If

                    Case DIR_LEFT

                        If GetPlayerX(Index) > 0 Then
                            x = GetPlayerX(Index) - 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                    Case DIR_RIGHT

                        If GetPlayerX(Index) < map(GetPlayerMap(Index)).MaxX Then
                            x = GetPlayerX(Index) + 1
                            y = GetPlayerY(Index)
                        Else
                            Exit Sub
                        End If

                End Select

                ' Check if a key exists
                If map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then

                    ' Check if the key they are using matches the map key
                    If GetPlayerInvItemNum(Index, InvNum) = map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                        TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                        TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                        Set buffer = New clsBuffer
                        buffer.WriteInteger SMapKey
                        buffer.WriteLong x
                        buffer.WriteLong y
                        buffer.WriteLong 1
                        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
                        Set buffer = Nothing
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                        
                        Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, x, y)

                        ' Check if we are supposed to take away the item
                        If map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                            Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                            Call PlayerMsg(Index, "The key is destroyed in the lock.", Yellow)
                        End If
                    End If
                End If

            Case ITEM_TYPE_SPELL
            
                For i = 1 To Stats.Stat_Count - 1
                    If GetPlayerStat(Index, i) < item(GetPlayerInvItemNum(Index, InvNum)).Stat_Req(i) Then
                        PlayerMsg Index, "You do not meet the stat requirements to use this item.", BrightRed
                        Exit Sub
                    End If
                Next
                
                ' Get the spell num
                n = item(GetPlayerInvItemNum(Index, InvNum)).Data1

                If n > 0 Then

                    ' Make sure they are the right class
                    If Spell(n).ClassReq = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                        ' Make sure they are the right level
                        i = Spell(n).LevelReq

                        If i <= GetPlayerLevel(Index) Then
                            i = FindOpenSpellSlot(Index)

                            ' Make sure they have an open spell slot
                            If i > 0 Then

                                ' Make sure they dont already have the spell
                                If Not HasSpell(Index, n) Then
                                    Call SetPlayerSpell(Index, i, n)
                                    Call SendAnimation(GetPlayerMap(Index), item(GetPlayerInvItemNum(Index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, Index)
                                    Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                    Call PlayerMsg(Index, "You study the spell carefully.", Yellow)
                                    Call PlayerMsg(Index, "You have learned a new spell!", White)
                                Else
                                    Call PlayerMsg(Index, "You have already learned this spell!", BrightRed)
                                End If

                            Else
                                Call PlayerMsg(Index, "You have learned all that you can learn!", BrightRed)
                            End If

                        Else
                            Call PlayerMsg(Index, "You must be level " & i & " to learn this spell.", White)
                        End If

                    Else
                        Call PlayerMsg(Index, "This spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", White)
                    End If

                Else
                    Call PlayerMsg(Index, "This scroll is not connected to a spell, please inform an admin!", White)
                End If
          Case ITEM_TYPE_POKEPOTION
        'For x = 1 To 6
        'If player(index).PokemonInstance(x).PokemonNumber > 0 Then
            'player(index).PokemonInstance(x).Hp = player(index).PokemonInstance(x).MaxHp
        'End If
        'Next
        'PlayerMsg index, "All your pokémon have been healed.", Green
        'SendPlayerPokemon index
        'Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 1)
        
        Case ITEM_TYPE_SCRIPT
        Call ItemCustomScript(Index, GetPlayerInvItemNum(Index, InvNum))
        If DoesItemTake(GetPlayerInvItemNum(Index, InvNum)) Then
        Else
        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 1)
        End If
        End Select

    End If
SendPlayerData Index
SendInventory Index
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Dim damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    
    ' can't attack whilst casting
    If TempPlayer(Index).SpellBuffer > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub

    ' Try to attack a player
    For i = 1 To MAX_PLAYERS
        TempIndex = i

        ' Make sure we dont try to attack ourselves
        If TempIndex <> Index Then

            ' Can we attack the player?
            If CanAttackPlayer(Index, TempIndex) Then
                If Not CanPlayerBlockHit(TempIndex) Then

                    ' Get the damage we can do
                    If Not CanPlayerCriticalHit(Index) Then
                        damage = GetPlayerDamage(Index) - GetPlayerProtection(TempIndex)
                    Else
                        n = GetPlayerDamage(Index)
                        damage = n + Int(Rnd * (n \ 2)) + 1 - GetPlayerProtection(TempIndex)
                        'Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                        'Call PlayerMsg(TempIndex, GetPlayerName(Index) & " swings with enormous might!", BrightCyan)
                        SendActionMsg GetPlayerMap(Index), "CRITICAL HIT!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
                    End If

                    Call AttackPlayer(Index, TempIndex, damage)
                Else
                    'Call PlayerMsg(Index, GetPlayerName(TempIndex) & "'s " & Trim$(Item(GetPlayerEquipment(TempIndex, Shield)).Name) & " has blocked your hit!", BrightCyan)
                    'Call PlayerMsg(TempIndex, "Your " & Trim$(Item(GetPlayerEquipment(TempIndex, Shield)).Name) & " has blocked " & GetPlayerName(Index) & "'s hit!", BrightCyan)
                    SendActionMsg GetPlayerMap(TempIndex), "BLOCK!", Pink, 1, (GetPlayerX(TempIndex) * 32), (GetPlayerY(TempIndex) * 32)
                End If

                Exit Sub
            End If
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS

        ' Can we attack the npc?
        If CanAttackNpc(Index, i) Then

            ' Get the damage we can do
            If Not CanPlayerCriticalHit(Index) Then
                damage = GetPlayerDamage(Index) - (NPC(MapNpc(GetPlayerMap(Index)).NPC(i).Num).Stat(Stats.endurance) \ 2)
            Else
                n = GetPlayerDamage(Index)
                damage = n + Int(Rnd * (n \ 2)) + 1 - (NPC(MapNpc(GetPlayerMap(Index)).NPC(i).Num).Stat(Stats.endurance) \ 2)
                'Call PlayerMsg(Index, "You feel a surge of energy upon swinging!", BrightCyan)
                SendActionMsg GetPlayerMap(Index), "CRITICAL HIT!", BrightCyan, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            End If

            If damage > 0 Then
                Call AttackNpc(Index, i, damage)
            Else
                Call PlayerMsg(Index, "Your attack does nothing.", BrightRed)
            End If

            Exit Sub
        End If

    Next

    ' Check tradeskills
    Select Case GetPlayerDir(Index)
        Case DIR_UP

            If GetPlayerY(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) - 1
        Case DIR_DOWN

            If GetPlayerY(Index) = map(GetPlayerMap(Index)).MaxY Then Exit Sub
            x = GetPlayerX(Index)
            y = GetPlayerY(Index) + 1
        Case DIR_LEFT

            If GetPlayerX(Index) = 0 Then Exit Sub
            x = GetPlayerX(Index) - 1
            y = GetPlayerY(Index)
        Case DIR_RIGHT

            If GetPlayerX(Index) = map(GetPlayerMap(Index)).MaxX Then Exit Sub
            x = GetPlayerX(Index) + 1
            y = GetPlayerY(Index)
    End Select
    
    'Check fishing
    If map(player(Index).map).Tile(player(Index).x, player(Index).y).Type = TILE_TYPE_CUSTOMSCRIPT Then
    If map(player(Index).map).Tile(player(Index).x, player(Index).y).Data1 = 2 Then
    PlayerFish Index
    End If
    End If
    
    CheckResource Index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim PointType As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > 3) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(Index) > 0 Then
        ' Take away a stat point
        Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.strength
                Call SetPlayerStat(Index, Stats.strength, GetPlayerStat(Index, Stats.strength) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 STR!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            Case Stats.endurance
                Call SetPlayerStat(Index, Stats.endurance, GetPlayerStat(Index, Stats.endurance) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 END!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            Case Stats.vitality
                Call SetPlayerStat(Index, Stats.vitality, GetPlayerStat(Index, Stats.vitality) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 VIT!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            Case Stats.intelligence
                Call SetPlayerStat(Index, Stats.intelligence, GetPlayerStat(Index, Stats.intelligence) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 INT!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            Case Stats.willpower
                Call SetPlayerStat(Index, Stats.willpower, GetPlayerStat(Index, Stats.willpower) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 WILL!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            Case Stats.spirit
                Call SetPlayerStat(Index, Stats.spirit, GetPlayerStat(Index, Stats.spirit) + 1)
                SendActionMsg GetPlayerMap(Index), "+1 SPR!", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
        End Select

    Else
        Exit Sub
    End If

    ' Send the update
    Call SendStats(Index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    i = FindPlayer(Name)

    If i > 0 Then
        Call PlayerMsg(Index, "Account: " & Trim$(player(i).Login) & ", Name: " & GetPlayerName(i), BrightGreen)

        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "-=- " & GetPlayerName(i) & " -=-", BrightGreen)

        End If

    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(Index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            'Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(Index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp to yourself!", White)
    End If

End Sub

Sub HandlePCScan(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim pn As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    pn = buffer.ReadString
    If IsPlaying(FindPlayer(pn)) Then
    SendPCScanRequest FindPlayer(pn)
    End If
   
   Set buffer = Nothing
    

End Sub


Sub HandlePCScanResult(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim pn As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim bfr As clsBuffer
    Set bfr = New clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    pn = buffer.ReadLong
    bfr.WriteByte SPCScan
    bfr.WriteLong pn
    For i = 1 To pn
    bfr.WriteString buffer.ReadString
    Next
    For i = 1 To MAX_PLAYERS
    If IsPlaying(i) Then
    If player(i).Access >= 3 Then
    SendDataTo i, bfr.ToArray
    End If
    End If
    Next
    Set bfr = Nothing
   Set buffer = Nothing
    

End Sub





Sub HandleWarpAdmin(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim x As Long
    Dim y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    If Not IsPlaying(Index) Then Exit Sub
    If player(Index).Access < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    Call PlayerWarp(Index, GetPlayerMap(Index), x, y)
   
   Set buffer = Nothing
    

End Sub


Sub HandleRequests(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Data1 As Long
    Dim Data2 As Long
    Dim Data3 As String
    Dim Data4 As String
    Dim rtype As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Data1 = buffer.ReadLong
    Data2 = buffer.ReadLong
    Data3 = buffer.ReadString
    Data4 = buffer.ReadString
    rtype = buffer.ReadString
   
   Set buffer = Nothing
   Select Case rtype
   Case "NPC"
   SendScript Index, ReadText("Data\NPCScripts\" & player(Index).map & "I" & Data1 & ".txt")
   Case "TRADE"
   If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub
   If TempPlayer(Index).isTrading = NO Then
   If TempPlayer(TempPlayer(Index).Target).isTrading = NO Then
   PlayerMsg TempPlayer(Index).Target, Trim$(player(Index).Name) & " wants to trade you!", Yellow
   'SendActionMsg player(index).map, "Hey " & Trim$(player(TempPlayer(index).Target).name) & " let's trade!", Yellow, ACTIONMSG_SCROLL, player(index).x * 32, player(index).y * 32
   'SendActionMsg TempPlayer(index).Target, "Hey " & Trim$(player(TempPlayer(index).Target).name) & " let's trade!", Yellow, ACTIONMSG_SCROLL, player(index).x * 32, player(index).y * 32
   SetTrade Index, player(TempPlayer(Index).Target).Name
   If Trim$(TempPlayer(TempPlayer(Index).Target).TradeName) = Trim$(player(Index).Name) Then
   'TRADEEEEEEEEEEEEEEEEE
   TempPlayer(Index).isTrading = YES
   TempPlayer(TempPlayer(Index).Target).isTrading = YES
   PlayerMsg Index, "Trade started!", BrightGreen
   PlayerMsg TempPlayer(Index).Target, "Trade started!", BrightGreen
   TempPlayer(Index).TradeItem = 0
   TempPlayer(Index).TradeItemVal = 0
   TempPlayer(Index).TradePoke = 0
   TempPlayer(Index).TradeLocked = 0
   TempPlayer(TempPlayer(Index).Target).TradeItem = 0
   TempPlayer(TempPlayer(Index).Target).TradeItemVal = 0
   TempPlayer(TempPlayer(Index).Target).TradePoke = 0
   TempPlayer(TempPlayer(Index).Target).TradeLocked = 0
   SendTradeStart Index
   SendTradeStart TempPlayer(Index).Target
   End If
   End If
   End If
   'check trade
   Case "TRADEUPDATE"
   SendTradeUpdate FindPlayer(Trim$(TempPlayer(Index).TradeName)), Data2, Data1, Data3
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeLocked = NO
   SendTradeLocked Index, NO
   SendTradeLocked FindPlayer(Trim$(TempPlayer(Index).TradeName)), YES
   Case "TRADESTOP"
   SendTradeStop Index
   If TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).isTrading = YES Then
   SendTradeStop FindPlayer(Trim$(TempPlayer(Index).TradeName))
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeName = ""
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).isTrading = NO
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradePoke = 0
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeItem = 0
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeItemVal = 0
   TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeLocked = NO
   End If
   TempPlayer(Index).TradeName = ""
   TempPlayer(Index).isTrading = NO
   TempPlayer(Index).TradePoke = 0
   TempPlayer(Index).TradeItem = 0
   TempPlayer(Index).TradeItemVal = 0
   TempPlayer(Index).TradeLocked = NO
   Case "TRADELOCK"
   If Data1 > 0 Then
   TempPlayer(Index).TradeItem = Data1
   Else
   TempPlayer(Index).TradeItem = 0
   End If
   If Data2 > 0 Then
   TempPlayer(Index).TradePoke = Data2
   Else
   TempPlayer(Index).TradePoke = 0
   End If
   If Val(Data3) > 0 Then
   TempPlayer(Index).TradeItemVal = Val(Data3)
   Else
   TempPlayer(Index).TradeItemVal = 1
   End If
   TempPlayer(Index).TradeLocked = YES
    SendTradeLocked Index, YES
   SendTradeLocked FindPlayer(Trim$(TempPlayer(Index).TradeName)), NO
   If TempPlayer(Index).TradeLocked = YES And TempPlayer(FindPlayer(Trim$(TempPlayer(Index).TradeName))).TradeLocked = YES Then
   DoTrade Index
   End If
   Case "FINDPOKEMON"
   PlayerMsg Index, Data1 & ": " & NumToName1(Val(Data1)), Yellow
   Case "FINDPOKEMONBYNAME"
   PlayerMsg Index, Data3 & ": " & NameToNum1(Data3), Yellow
   Case "TRYLOGIN"
   JoinGame (Index)
   Case "TRYEVOLVE"
   Call TryPokemonEvolution(Index, Data1)
   Case "PEV"
   Call EvolvePokemon(Index, Data1)
   Case "TRYLM"
   CheckCustomLearnMove Index, Data1, Data2
   Case "TRAVEL"
   CheckTravel Index, Data1
   Case "GPOKE"
   If GetPlayerAccess(Index) >= 4 Then
   GivePokemon Index, Data1
   End If
   Case "GITEM"
   If GetPlayerAccess(Index) >= 4 Then
   GiveItem Index, Data1, 1
   End If
   Case "EMOTE"
   GlobalMsg Data3, BrightRed
   Case "PCSCAN"
   SendPCScan FindPlayer(Trim$(Data4)), Trim$(Data3)
   Case "PCSCANRESULT"
   SendPCScanResultToAdmin FindPlayer(Data4), Data3
   Case "USEITEMONPOKEMON"
   UseItemOnPokemon Index, Data1, Data2
   Case "PRIVATEM"
   If IsPlaying(FindPlayer(Trim$(Data4))) Then
   PlayerMsg FindPlayer(Trim$(Data4)), "From " & Trim$(GetPlayerName(Index)) & ": " & Trim$(Data3), Pink
   PlayerMsg Index, "To " & Trim$(Data4) & ": " & Trim$(Data3), Pink
   Else
   PlayerMsg Index, "That player is offline!", BrightRed
   End If
   Case "RADIOPLAY"
   SendRadio 1, Data3
   PlayerMsg Index, "[Jigglypuff Radio] Song has been changed!", BrightCyan
   Case "MUTE"
   If IsPlaying(FindPlayer(Data3)) Then
   If IsPlaying(Index) Then
   If GetPlayerAccess(Index) >= GetPlayerAccess(FindPlayer(Data3)) Then
   If isPlayerMuted(FindPlayer(Data3)) Then
   GlobalMsg "[GM] " & GetPlayerName(FindPlayer(Data3)) & " has been unmuted by " & GetPlayerName(Index) & "!", White
   Else
   GlobalMsg "[GM] " & GetPlayerName(FindPlayer(Data3)) & " has been muted by " & GetPlayerName(Index) & "!", White
   End If
   MutePlayer FindPlayer(Data3)
   End If
   End If
   Else
   PlayerMsg Index, "Player is not online", BrightRed
   End If
   Case "BAGUPDATE"
   SendInventory Index
   Case "WHOSDATPOKE"
   If UCase(Data3) = UCase(WhosDatPokemon) Then
    isWhosOn = False
   GlobalMsg GetPlayerName(Index) & " has guessed the pokemon! Its " & UCase(WhosDatPokemon) & "!", Yellow
   GlobalMsg GetPlayerName(Index) & " has won x" & WhosRewardItemVal & " " & Trim$(item(WhosRewardItem).Name) & "!", Yellow
   GiveItem Index, WhosRewardItem, WhosRewardItemVal
   WhosRewardItem = 0
   WhosRewardItemVal = 0
   WhosDatPokemon = ""
   SendCloseWhos 1, 1
   Else
   GlobalMsg GetPlayerName(Index) & " thinks its " & UCase(Data3) & "!", White
   End If
   Case "PPIC"
   SetPlayerProfilePicture Index, Data3
   Case "MAKECREW"
   If GetPlayerInvItemValue(Index, GetItemSlot(Index, 1)) >= 25000 Then
   If GetPlayerCrew(Index) = "" Then
   Call MakeCrew(Index, Data3)
   TakeItem Index, 1, 25000
   End If
   End If
   
   Case "TPREMOVE"
   If TempPlayer(Index).isInTPRemoval = True Then
   Call RemoveTP(Index, Data1, Data2)
   End If
   
   Case "BATTLE"
   If TempPlayer(Index).TargetType <> TARGET_TYPE_PLAYER Then Exit Sub
   If TempPlayer(Index).isInPVP = False Then
   If TempPlayer(TempPlayer(Index).Target).isInPVP = False Then
   PlayerMsg TempPlayer(Index).Target, Trim$(player(Index).Name) & " wants to battle you!", Yellow
   TempPlayer(Index).PVPEnemy = player(TempPlayer(Index).Target).Name
   If Trim$(TempPlayer(TempPlayer(Index).Target).PVPEnemy) = Trim$(player(Index).Name) Then
   'BATTLE
   TempPlayer(Index).isInPVP = True
   TempPlayer(TempPlayer(Index).Target).isInPVP = True
   PlayerMsg Index, "Battle started!", BrightGreen
   PlayerMsg TempPlayer(Index).Target, "Battle started!", BrightGreen
   SendPVPCommand Index, "PVP"
   SendPVPCommand TempPlayer(Index).Target, "PVP"
    StartPVPBattle Index
   End If
   End If
   End If
   
   Case "CREW"
   If DoesCrewExist(GetPlayerCrew(Index)) Then
   Call SendCrewData(Index, GetPlayerCrew(Index))
   End If
   
   Case "CREWPICTURE"
   If DoesCrewExist(GetPlayerCrew(Index)) Then
   If Trim$(GetPlayerName(Index)) = Trim$(GetCrewLeaderName(GetPlayerCrew(Index))) Then
   Call PutVar(App.Path & "\Data\crews\" & GetPlayerCrew(Index) & ".ini", "DATA", "Picture", Data3)
   PlayerMsg Index, "Clan image set!", Yellow
   SendCrewData Index, GetPlayerCrew(Index)
   Else
   PlayerMsg Index, "You are not the leader of the clan!", BrightRed
   End If
   End If
   
   Case "CREWDELETE"
   
    If DoesCrewExist(GetPlayerCrew(Index)) Then
   If Trim$(GetPlayerName(Index)) = Trim$(GetCrewLeaderName(GetPlayerCrew(Index))) Then
   Call PutVar(App.Path & "\Data\crews\" & GetPlayerCrew(Index) & ".ini", "DATA", "Picture", Data3)
   DeleteCrew GetPlayerCrew(Index)
   Else
   PlayerMsg Index, "You are not the leader of the clan!", BrightRed
   End If
   End If
   
   
   Case "JOURNAL"
   SendJournal Index, FindPlayer(Trim$(Data3))
   
   Case "SAVEJOURNAL"
   WriteText App.Path & "\Data\journals\" & GetPlayerName(Index) & ".txt", Data3
   PlayerMsg Index, "Journal saved!", Yellow
   
   Case "CLANINVITE"
   Dim clanTarget As Long
   clanTarget = FindPlayer(Data3)
   TempPlayer(clanTarget).clanInvite = True
   TempPlayer(clanTarget).clanInviteIndex = Index
   SendClanInvite clanTarget, Index, GetPlayerCrew(Index)
   
   Case "CLANRESPOND"
   If TempPlayer(Index).clanInvite = True Then
   If Data3 = "YES" Then
   AddToCrew Index, GetPlayerCrew(TempPlayer(Index).clanInviteIndex)
   TempPlayer(Index).clanInvite = False
   TempPlayer(Index).clanInviteIndex = 0
   Else
   PlayerMsg GetPlayerCrew(TempPlayer(Index).clanInviteIndex), GetPlayerName(Index) & " refused to join the crew", Yellow
   TempPlayer(Index).clanInvite = False
   TempPlayer(Index).clanInviteIndex = 0
   End If
   End If
   
   Case "CREWKICK"
   If Data1 < 1 Or Data1 > 50 Then
   PlayerMsg Index, "You cant kick yourself!", Yellow
   Else
   If Trim$(GetPlayerName(Index)) = Trim$(GetCrewLeaderName(GetPlayerCrew(Index))) Then
   If GetPlayerCrewByName(GetCrewPlayerName(GetPlayerCrew(Index), Data1)) = GetPlayerCrew(Index) Then
   ClanMsg GetPlayerCrew(Index), GetCrewPlayerName(GetPlayerCrew(Index), Data1) & " has been kicked from clan!"
   Call RemoveMemberFromCrew(GetPlayerCrew(Index), Data1)
   End If
   End If
   End If
   
   Case "CREWLEAVE"
   If Trim$(GetPlayerName(Index)) = Trim$(GetCrewLeaderName(GetPlayerCrew(Index))) Then
   PlayerMsg Index, "You cant leave clan if you are leader!In order to leave you must delete it!", Yellow
   Else
   If GetPlayerCrew(Index) <> "" Then
   ClanMsg GetPlayerCrew(Index), GetPlayerName(Index) & " has left the clan!"
   Call RemoveMemberFromCrew(GetPlayerCrew(Index), GetPlayerCrewSpot(Index, GetPlayerCrew(Index)))
   End If
   End If
   
   Case "VISIBLE"
   If GetPlayerAccess(Index) >= ADMIN_MAPPER Then
   If TempPlayer(Index).notVisible = True Then
   TempPlayer(Index).notVisible = False
   PlayerMsg Index, "[ADMIN] You are now visible!", BrightCyan
   SendPlayerData Index
   Else
   TempPlayer(Index).notVisible = True
   PlayerMsg Index, "[ADMIN] You are now invisible!", BrightCyan
   SendPlayerData Index
   End If
   End If
   
   Case "CLANNEWS"
   If Trim$(GetPlayerName(Index)) = Trim$(GetCrewLeaderName(GetPlayerCrew(Index))) Then
   WriteText App.Path & "\Data\clanNews\" & GetPlayerCrew(Index) & ".txt", Data3
   PlayerMsg Index, "Clan news updated!", Yellow
   SendCrewData Index, GetPlayerCrew(Index)
   End If
   
   Case "OPENNEWS"
   SendNews Index
   
   Case "HEAL"
   If GetPlayerAccess(Index) >= 3 Then
   HealPokemons (Index)
   End If
   
   
   Case "PROFILE"
   SendProfile Index
   
   Case "DIALOGTRIGGER"
   If TempPlayer(Index).hasDialogTrigger Then
   'PlayerMsg index, "YY", Yellow
   Select Case TempPlayer(Index).dialogTriggerData1
   Case DIALOG_NPCBATTLE
   StartNPCBattle Index, TempPlayer(Index).dialogTriggerData3
   TempPlayer(Index).dialogTriggerData1 = 0
   TempPlayer(Index).dialogTriggerData2 = 0
   TempPlayer(Index).dialogTriggerData3 = ""
   TempPlayer(Index).hasDialogTrigger = False
    'PlayerMsg index, "XX", Yellow
    Case DIALOG_GIVEITEM
    GiveItem Index, TempPlayer(Index).dialogTriggerData2, Val(TempPlayer(Index).dialogTriggerData3)
    TempPlayer(Index).dialogTriggerData1 = 0
   TempPlayer(Index).dialogTriggerData2 = 0
   TempPlayer(Index).dialogTriggerData3 = ""
   TempPlayer(Index).hasDialogTrigger = False
   End Select
   End If
   
   Case "EGG"
   If DoesPlayerHaveEgg(Index) Then
   SaveEggFromTemp Index
   SendEgg Index
   Else
   PlayerMsg Index, "You dont have an egg equipped!", Yellow
   End If
   
   Case "HATCHEGG"
   If DoesPlayerHaveEgg(Index) Then HatchEgg (Index)
   
   
   Case "BIKE"
   If DoesPlayerHaveBike(Index) Then
   UseBike Index
   Else
   PlayerMsg Index, "You don't own a bike!", Yellow
   End If
   
   Case "MARKET"
   Set buffer = New clsBuffer
                buffer.WriteInteger SOpenShop ' send packet opening the shop
                buffer.WriteLong 10
                SendDataTo Index, buffer.ToArray()
                Set buffer = Nothing
                TempPlayer(Index).InShop = 10 ' stops movement and the like
                
                
                Case "CLANMSG"
                If Trim$(GetPlayerCrew(Index)) <> "" Then
                Call ClanMsg(GetPlayerCrew(Index), GetPlayerName(Index) & ":" & Data3)
                End If
                
   
   '------------END------------------------
   End Select

End Sub


Sub HandleSetMood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim x As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong
   Set buffer = Nothing
    Call SetPlayerMood(Index, x)
    SendPlayerData Index
End Sub


Sub HandleSetMapMusic(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim x As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadString
   Set buffer = Nothing
    If x = "" Then Exit Sub
    Call PutVar(App.Path & "\Data\MapData\" & GetPlayerMap(Index) & ".ini", "DATA", "Music", x)
    PlayerMsg Index, "Music " & x & " set to map " & GetPlayerMap(Index), Yellow
    SendMusicToMap GetPlayerMap(Index)
End Sub


' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(Index) & ".", BrightBlue)
            Call PlayerMsg(Index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(Index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(Index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot warp yourself to yourself!", White)
    End If

End Sub


Sub HandleDepositPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long
    Dim pcslot As Long
    
    For i = 1 To MAX_INV
    If GetPlayerInvItemNum(Index, i) = 1 Then
    pcslot = i
    Exit For
    End If
    Next
    
    If pcslot <= 0 Then
    PlayerMsg Index, "You dont have enough PokeCoins!", BrightRed
    Else
    '
    If GetPlayerInvItemValue(Index, pcslot) >= 500 Then
    SetPlayerInvItemValue Index, pcslot, GetPlayerInvItemValue(Index, pcslot) - 500
    player(Index).StoredPC = player(Index).StoredPC + 500
    SendUpdateBank (Index)
    SendPlayerData (Index)
    SendInventoryUpdate Index, pcslot
    Else
    PlayerMsg Index, "You don't have enough PokeCoins!", BrightRed
    End If
    '
    End If
    SendUpdateBank (Index)
End Sub

Sub HandleWithdrawPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim i As Long
Dim pcslot As Long
Dim invslot As Long
Dim a As Long
For i = 1 To MAX_INV
If GetPlayerInvItemNum(Index, i) = 1 Then
pcslot = i
Exit For
End If
Next



If player(Index).StoredPC >= 500 Then
If pcslot <= 0 Then
pcslot = FindOpenInvSlot(Index, 1)
SetPlayerInvItemNum Index, pcslot, 1
End If
SetPlayerInvItemValue Index, pcslot, GetPlayerInvItemValue(Index, pcslot) + 500
player(Index).StoredPC = player(Index).StoredPC - 500
SendUpdateBank (Index)
SendPlayerData (Index)
SendInventoryUpdate Index, pcslot
Else
PlayerMsg Index, "You dont have enough stored PokeCoins!", BrightRed
End If
SendUpdateBank (Index)
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(Index, n, GetPlayerX(Index), GetPlayerY(Index))
    Call PlayerMsg(Index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(Index) & " warped to map #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    Call SetPlayerSprite(Index, n)
    Call SendPlayerData(Index)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long
    Dim n As Long
    Call PlayerMsg(Index, "-=- Stats for " & GetPlayerName(Index) & " -=-", White)
    Call PlayerMsg(Index, "Level: " & GetPlayerLevel(Index) & "  Exp: " & GetPlayerExp(Index) & "/" & GetPlayerNextLevel(Index), White)
    Call PlayerMsg(Index, "HP: " & GetPlayerVital(Index, Vitals.hp) & "/" & GetPlayerMaxVital(Index, Vitals.hp) & "  MP: " & GetPlayerVital(Index, Vitals.mp) & "/" & GetPlayerMaxVital(Index, Vitals.mp) & "  SP: " & GetPlayerVital(Index, Vitals.SP) & "/" & GetPlayerMaxVital(Index, Vitals.SP), White)
    Call PlayerMsg(Index, "STR: " & GetPlayerStat(Index, Stats.strength) & "  DEF: " & GetPlayerStat(Index, Stats.endurance) & "  MAGI: " & GetPlayerStat(Index, Stats.intelligence) & "  Speed: " & GetPlayerStat(Index, Stats.spirit), White)
    n = (GetPlayerStat(Index, Stats.strength) \ 2) + (GetPlayerLevel(Index) \ 2)
    i = (GetPlayerStat(Index, Stats.endurance) \ 2) + (GetPlayerLevel(Index) \ 2)

    If n > 100 Then n = 100
    If i > 100 Then i = 100
    Call PlayerMsg(Index, "Critical Hit Chance: " & n & "%, Block Chance: " & i & "%", White)
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim i As Long
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 0
    mapnum = GetPlayerMap(Index)
    i = map(mapnum).Revision + 1
    Call ClearMap(mapnum)
    map(mapnum).Name = buffer.ReadString 'Parse(n + 1)
    map(mapnum).Revision = i
    map(mapnum).Moral = buffer.ReadLong 'CByte(Parse(n + 2))
    map(mapnum).Tileset = buffer.ReadLong 'CInt(Parse(n + 3))
    map(mapnum).Up = buffer.ReadLong 'CInt(Parse(n + 4))
    map(mapnum).Down = buffer.ReadLong 'CInt(Parse(n + 5))
    map(mapnum).Left = buffer.ReadLong 'CInt(Parse(n + 6))
    map(mapnum).Right = buffer.ReadLong 'CInt(Parse(n + 7))
    map(mapnum).Music = buffer.ReadLong 'CByte(Parse(n + 8))
    map(mapnum).BootMap = buffer.ReadLong 'CByte(Parse(n + 9))
    map(mapnum).BootX = buffer.ReadLong 'CByte(Parse(n + 10))
    map(mapnum).BootY = buffer.ReadLong 'CByte(Parse(n + 11))
    map(mapnum).MaxX = buffer.ReadLong 'CByte(Parse(n + 13))
    map(mapnum).MaxY = buffer.ReadLong 'CByte(Parse(n + 14))
    ReDim map(mapnum).Tile(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)
    n = n + 15

    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                map(mapnum).Tile(x, y).Layer(i).x = buffer.ReadByte
                map(mapnum).Tile(x, y).Layer(i).y = buffer.ReadByte
                map(mapnum).Tile(x, y).Layer(i).Tileset = buffer.ReadByte
            Next
            map(mapnum).Tile(x, y).Type = buffer.ReadLong 'CByte(Parse(n + 6))
            map(mapnum).Tile(x, y).Data1 = buffer.ReadLong 'CInt(Parse(n + 7))
            map(mapnum).Tile(x, y).Data2 = buffer.ReadLong 'CInt(Parse(n + 8))
            map(mapnum).Tile(x, y).Data3 = buffer.ReadLong 'CInt(Parse(n + 9))
            n = n + 10
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        map(mapnum).NPC(x) = buffer.ReadLong 'CByte(Parse(n))
        Call ClearMapNpc(x, mapnum)
        n = n + 1
    Next

    For x = 1 To MAX_MAP_POKEMONS
    map(mapnum).Pokemon(x).PokemonNumber = buffer.ReadLong
    map(mapnum).Pokemon(x).LevelFrom = buffer.ReadLong
    map(mapnum).Pokemon(x).LevelTo = buffer.ReadLong
    map(mapnum).Pokemon(x).Custom = buffer.ReadLong
    map(mapnum).Pokemon(x).atk = buffer.ReadLong
    map(mapnum).Pokemon(x).def = buffer.ReadLong
    map(mapnum).Pokemon(x).spatk = buffer.ReadLong
    map(mapnum).Pokemon(x).spdef = buffer.ReadLong
    map(mapnum).Pokemon(x).spd = buffer.ReadLong
    map(mapnum).Pokemon(x).hp = buffer.ReadLong
    map(mapnum).Pokemon(x).Chance = buffer.ReadLong
    Next
    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))
    ' Save the map
    Call SaveMap(mapnum)
    Call MapCache_Create(mapnum)
    Call ClearTempTile(mapnum)
    Call CacheResources(mapnum)

    ' Refresh map for everyone online
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = mapnum Then
            Call PlayerWarp(i, mapnum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim s As String
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Get yes/no value
    s = buffer.ReadLong 'Parse(1)
    Set buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)

    'send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
        SendResourceCacheTo Index, i
    Next

    TempPlayer(Index).GettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteInteger SMapDone
    SendDataTo Index, buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Call PlayerMapGetItem(Index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim InvNum As Long
    Dim amount As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvNum = buffer.ReadLong 'CLng(Parse(1))
    amount = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    If GetPlayerInvItemNum(Index, InvNum) < 1 Or GetPlayerInvItemNum(Index, InvNum) > MAX_ITEMS Then Exit Sub
    If item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
        If amount < 1 Or amount > GetPlayerInvItemValue(Index, InvNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(Index, InvNum, amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
        Call ClearMapItem(i, GetPlayerMap(Index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(Index))
    Next

    CacheResources GetPlayerMap(Index)
    Call PlayerMsg(Index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Free Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If LenB(Trim$(map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(Index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(Index) & "!", White)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot kick yourself!", White)
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(Index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim FileName As String
    Dim file As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    FileName = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open FileName For Output As #F
        Close #F
    End If

    Kill FileName
    Call PlayerMsg(Index, "Ban list destroyed.", White)
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If n <> Index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(Index) Then
                Call BanIndex(n, Index)
            Else
                Call PlayerMsg(Index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "You cannot ban yourself!", White)
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SEditMap
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SItemEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SAnimationEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = buffer.ReadLong 'CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(Index) & " saved Animation #" & n & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SNpcEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim NpcNum As Long
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NpcNum = buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(NpcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(NpcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(Index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SResourceEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(Index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Pokemon packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SPokemonEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub


Sub HandleRequestEditMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
Dim i As Long
   
    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If
 For i = 1 To MAX_MOVES
    SendUpdateMoveTo Index, i
    Next
    Set buffer = New clsBuffer
    buffer.WriteInteger SMovesEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
    
    
End Sub


' :::::::::::::::::::::
' :: Save Pokemon packet ::
' :::::::::::::::::::::
Private Sub HandleSavePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim pokemonnum As Long
    Dim buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    pokemonnum = buffer.ReadLong

    ' Prevent hacking
    If pokemonnum < 0 Or pokemonnum > MAX_POKEMONS Then
        Exit Sub
    End If

    PokemonSize = LenB(Pokemon(pokemonnum))
    ReDim PokemonData(PokemonSize)
    PokemonData = buffer.ReadBytes(PokemonSize - 1)
    CopyMemory ByVal VarPtr(Pokemon(pokemonnum)), ByVal VarPtr(PokemonData(0)), PokemonSize
    ' Save it
    Call SendUpdatePokemonToAll(pokemonnum)
    Call SavePokemon(pokemonnum)
    Call AddLog(GetPlayerName(Index) & " saved Pokemon #" & pokemonnum & ".", ADMIN_LOG)
End Sub

Private Sub HandleSaveMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim moveNum As Long
    Dim buffer As clsBuffer
    Dim Movesize As Long
    Dim moveData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    moveNum = buffer.ReadLong

    ' Prevent hacking
    If moveNum < 0 Or moveNum > 500 Then
        Exit Sub
    End If

    Movesize = LenB(PokemonMove(moveNum))
    ReDim moveData(Movesize - 1)
    moveData = buffer.ReadBytes(Movesize)
    CopyMemory ByVal VarPtr(PokemonMove(moveNum)), ByVal VarPtr(moveData(0)), Movesize
    ' Save it
    Call SendUpdateMoveToAll(moveNum)
    Call SaveMove(moveNum)
    Call AddLog(GetPlayerName(Index) & " saved Move #" & moveNum & ".", ADMIN_LOG)
End Sub


' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SShopEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim ShopNum As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ShopNum = buffer.ReadLong

    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteInteger SSpellEditor
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim spellnum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    spellnum = buffer.ReadLong

    ' Prevent hacking
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)
    Call AddLog(GetPlayerName(Index) & " saved Spell #" & spellnum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The access
    i = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(Index) Then
                Call PlayerMsg(Index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(Index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(Index, "Invalid access level.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Call SendWhosOnline(Index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(buffer.ReadString) 'Parse(1))
    SaveOptions
    Set buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong 'CLng(Parse(1))
    y = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Prevent subscript out of range
    If x < 0 Or x > map(GetPlayerMap(Index)).MaxX Or y < 0 Or y > map(GetPlayerMap(Index)).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            If GetPlayerMap(Index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then

                        ' Consider the player
                        If i <> Index Then
                            If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                                'Call PlayerMsg(index, "You wouldn't stand a chance.", BrightRed)
                            Else

                                If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                                    'Call PlayerMsg(index, "This one seems to have an advantage over you.", Yellow)
                                Else

                                    If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                                        'Call PlayerMsg(index, "This would be an even fight.", White)
                                    Else

                                        If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                            'Call PlayerMsg(index, "You could slaughter that player.", BrightBlue)
                                        Else

                                            If GetPlayerLevel(Index) > GetPlayerLevel(i) Then
                                                'Call PlayerMsg(index, "You would have an advantage over that player.", Yellow)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Change target
                        TempPlayer(Index).Target = i
                        TempPlayer(Index).TargetType = TARGET_TYPE_PLAYER
                        'Call PlayerMsg(index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                        If TempPlayer(i).notVisible And i <> Index Then Exit Sub
                        SendTrainerCard Index, i
                        Exit Sub
                    End If
                End If
            End If
        End If

    Next

    ' Check for an item
    For i = 1 To MAX_MAP_ITEMS

        If MapItem(GetPlayerMap(Index), i).Num > 0 Then
            If MapItem(GetPlayerMap(Index), i).x = x Then
                If MapItem(GetPlayerMap(Index), i).y = y Then
                    Call PlayerMsg(Index, "You see " & CheckGrammar(Trim$(item(MapItem(GetPlayerMap(Index), i).Num).Name)) & ".", Yellow)
                    Exit Sub
                End If
            End If
        End If

    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS

        If MapNpc(GetPlayerMap(Index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(Index)).NPC(i).x = x Then
                If MapNpc(GetPlayerMap(Index)).NPC(i).y = y Then
                    ' Change target
                    TempPlayer(Index).Target = i
                    TempPlayer(Index).TargetType = TARGET_TYPE_NPC
                    'Call PlayerMsg(index, "Your target is now " & CheckGrammar(Trim$(Npc(MapNpc(GetPlayerMap(index)).Npc(i).Num).Name)) & ".", Yellow)
                    Exit Sub
                End If
            End If
        End If

    Next

End Sub

' ::::::::::::::::::
' :: Party packet ::
' ::::::::::::::::::
Sub HandleParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Exit Sub
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    ' Prevent partying with self
    If n = Index Then
        Exit Sub
    End If

    ' Check for a previous party and if so drop it
    If TempPlayer(Index).InParty = YES Then
        Call PlayerMsg(Index, "You are already in a party!", BrightRed)
        Exit Sub
    End If

    If n > 0 Then

        ' Check if its an admin
        If GetPlayerAccess(Index) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "You can't join a party, you are an admin!", BrightBlue)
            Exit Sub
        End If

        If GetPlayerAccess(n) > ADMIN_MONITOR Then
            Call PlayerMsg(Index, "Admins cannot join parties!", BrightBlue)
            Exit Sub
        End If

        ' Make sure they are in right level range
        If GetPlayerLevel(Index) + 5 < GetPlayerLevel(n) Or GetPlayerLevel(Index) - 5 > GetPlayerLevel(n) Then
            Call PlayerMsg(Index, "There is more then a 5 level gap between you two, party failed.", BrightRed)
            Exit Sub
        End If

        ' Check to see if player is already in a party
        If TempPlayer(n).InParty = NO Then
            Call PlayerMsg(Index, "Party request has been sent to " & GetPlayerName(n) & ".", BrightBlue)
            Call PlayerMsg(n, GetPlayerName(Index) & " wants you to join their party.  Type /join to join, or /leave to decline.", BrightBlue)
            TempPlayer(Index).PartyStarter = YES
            TempPlayer(Index).PartyPlayer = n
            TempPlayer(n).PartyPlayer = Index
        Else
            Call PlayerMsg(Index, "Player is already in a party!", BrightRed)
        End If

    Else
        Call PlayerMsg(Index, "Player is not online.", White)
    End If

End Sub

' :::::::::::::::::::::::
' :: Join party packet ::
' :::::::::::::::::::::::
Sub HandleJoinParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    n = TempPlayer(Index).PartyPlayer

    If n > 0 Then

        ' Check to make sure they aren't the starter
        If TempPlayer(Index).PartyStarter = NO Then

            ' Check to make sure that each of there party players match
            If TempPlayer(n).PartyPlayer = Index Then
                Call PlayerMsg(Index, "You have joined " & GetPlayerName(n) & "'s party!", BrightGreen)
                Call PlayerMsg(n, GetPlayerName(Index) & " has joined your party!", BrightGreen)
                TempPlayer(Index).InParty = YES
                TempPlayer(n).InParty = YES
            Else
                Call PlayerMsg(Index, "Party failed.", BrightRed)
            End If

        Else
            Call PlayerMsg(Index, "You have not been invited to join a party!", BrightRed)
        End If

    Else
        Call PlayerMsg(Index, "You have not been invited into a party!", BrightRed)
    End If

End Sub

' ::::::::::::::::::::::::
' :: Leave party packet ::
' ::::::::::::::::::::::::
Sub HandleLeaveParty(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    n = TempPlayer(Index).PartyPlayer

    If n > 0 Then
        If TempPlayer(Index).InParty = YES Then
            Call PlayerMsg(Index, "You have left the party.", BrightBlue)
            Call PlayerMsg(n, GetPlayerName(Index) & " has left the party.", BrightBlue)
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        Else
            Call PlayerMsg(Index, "Declined party request.", BrightGreen)
            Call PlayerMsg(n, GetPlayerName(Index) & " declined your request.", BrightGreen)
            TempPlayer(Index).PartyPlayer = 0
            TempPlayer(Index).PartyStarter = NO
            TempPlayer(Index).InParty = NO
            TempPlayer(n).PartyPlayer = 0
            TempPlayer(n).PartyStarter = NO
            TempPlayer(n).InParty = NO
        End If

    Else
        Call PlayerMsg(Index, "You are not in a party!", BrightRed)
    End If

End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Call SendPlayerSpells(Index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Spell slot
    n = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(Index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Call CloseSocket(Index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Dim OldSlot As Integer, NewSlot As Integer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    OldSlot = buffer.ReadInteger
    NewSlot = buffer.ReadInteger
    Set buffer = Nothing
    PlayerSwitchInvSlots Index, OldSlot, NewSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteInteger SSendPing
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleUnequip(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerUnequipItem Index, buffer.ReadLong
    Set buffer = Nothing
    SendPlayerData Index
    SendInventory Index
End Sub

Sub HandleRequestPlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendPlayerData Index
End Sub

Sub HandleRequestItems(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendItems Index
End Sub

Sub HandleRequestAnimations(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendAnimations Index
End Sub

Sub HandleRequestNPCS(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendNpcs Index
End Sub

Sub HandleRequestResources(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendResources Index
End Sub

Sub HandleRequestPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendPokemon Index
End Sub

Sub HandleRequestMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendMove Index
End Sub



Sub HandleRequestSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendSpells Index
End Sub

Sub HandleRequestShops(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SendShops Index
End Sub

Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index)
    Set buffer = Nothing
End Sub

Sub HandleTrainStat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim tmpStat As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' check points
    If GetPlayerPOINTS(Index) = 0 Then Exit Sub
    
    ' stat
    tmpStat = buffer.ReadByte
    
    ' increment stat
    SetPlayerStat Index, tmpStat, GetPlayerRawStat(Index, tmpStat) + 1
    
    ' decrement points
    SetPlayerPOINTS Index, GetPlayerPOINTS(Index) - 1
    
    ' send player new data
    SendPlayerData Index
    
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    SetPlayerExp Index, GetPlayerNextLevel(Index)
    CheckPlayerLevelUp Index
End Sub

Sub HandleForgetSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim spellslot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellslot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(Index).SpellCD(spellslot) > 0 Then
        PlayerMsg Index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(Index).SpellBuffer = spellslot Then
        PlayerMsg Index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    player(Index).Spell(spellslot) = 0
    SendPlayerSpells Index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    TempPlayer(Index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim ShopNum As Long
    Dim itemamount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    ShopNum = TempPlayer(Index).InShop
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(ShopNum).TradeItem(shopslot)
        ' check trade exists
        If .item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(Index, .costitem)
        If itemamount = 0 Or itemamount < .costvalue Then
            PlayerMsg Index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction Index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeItem Index, .costitem, .costvalue
        GiveItem Index, .item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Successful.", BrightGreen
    SendInventory Index
    SendResetShop Index
    ResetShopAction Index
    
    Set buffer = Nothing
End Sub

Sub HandleDepositPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim pokeSlot As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
pokeSlot = buffer.ReadLong
Set buffer = Nothing

'Deposit pokemon

DepositPokemon Index, pokeSlot
End Sub

Sub HandleSetAsLeader(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim slot As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
slot = buffer.ReadLong
Set buffer = Nothing
SetAsLeader Index, slot
End Sub

Sub HandleAddTP(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim Stat As Long
Dim pokeSlot As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
Stat = buffer.ReadLong
pokeSlot = buffer.ReadLong
Set buffer = Nothing

'AddTP

AddTP Index, Stat, pokeSlot
End Sub

Sub HandleRosterRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
SendPlayerPokemon Index
sendopenroster Index
'DONE!
End Sub

Sub HandleWithdrawPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim storageslot As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
storageslot = buffer.ReadLong
Set buffer = Nothing

'Withdraw
WithdrawPokemon Index, storageslot
End Sub


Sub HandleMapNPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim NpcNum As Long
Dim script As String
Dim y As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
NpcNum = buffer.ReadLong
script = buffer.ReadString
Set buffer = Nothing
WriteText App.Path & "\Data\NPCScripts\" & player(Index).map & "I" & NpcNum & ".txt", Trim$(script)
'MsgBox Script
PlayerMsg Index, "Script edited for npc: " & player(Index).map & " - MAP , " & NpcNum & " - NPC", Yellow
SendScript Index, ReadText("Data\NPCScripts\" & player(Index).map & "I" & NpcNum & ".txt")
End Sub


Sub HandleRemoveStoragePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim buffer As clsBuffer
Dim storageslot As Long
Set buffer = New clsBuffer
buffer.WriteBytes Data()
storageslot = buffer.ReadLong
Set buffer = Nothing

'Remove pokemon
RemoveStoragePokemon Index, storageslot
End Sub


Sub HandleSellItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim invslot As Long
    Dim itemNum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invslot = buffer.ReadLong
    
    ' if invalid, exit out
    If invslot < 1 Or invslot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(Index, invslot) < 1 Or GetPlayerInvItemNum(Index, invslot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemNum = GetPlayerInvItemNum(Index, invslot)
    
    ' work out price
    multiplier = Shop(TempPlayer(Index).InShop).BuyRate / 100
    price = item(itemNum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg Index, "The shop doesn't want that item.", BrightRed
        ResetShopAction Index
        Exit Sub
    End If

    ' take item and give gold
    TakeItem Index, itemNum, 1
    GiveItem Index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg Index, "Sell successful.", BrightGreen
    SendInventory Index
    SendResetShop Index
    ResetShopAction Index
    
    Set buffer = Nothing
End Sub

Sub HandleBattleCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim buffer As clsBuffer
    Dim command As Byte
    Dim damage As Long
    Dim pokeSlot As Long
    Dim move As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    command = buffer.ReadByte
    pokeSlot = buffer.ReadLong
    move = buffer.ReadLong
    BattleCommand Index, command, pokeSlot, move
    Set buffer = Nothing
End Sub

Public Function NumToName1(ByVal Num As Long) As String
On Error Resume Next
If Num < 1 Or Num > MAX_POKEMONS Then Exit Function
Dim str As String
str = Num
NumToName1 = GetVar(App.Path & "\Data\Pokemon Data\Nums_Names.ini", "DATA", str)
End Function

Public Function NameToNum1(ByVal Name As String) As Long
On Error Resume Next
If Name = vbNullString Or Name = "" Then Exit Function
NameToNum1 = Val(GetVar(App.Path & "\Data\Pokemon Data\Names_Nums.ini", "DATA", Name))
End Function

Sub HandleLearnMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long
    Dim a As Long
    Dim canLearn As Boolean
    Dim buffer As clsBuffer
    Dim moveSlot As Long
    Dim pokeSlot As Long
    Dim move As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    pokeSlot = buffer.ReadLong
    moveSlot = buffer.ReadLong
    move = buffer.ReadLong
    Set buffer = Nothing
    If Not TempPlayer(Index).LearnMovePokemonName = Trim$(Pokemon(player(Index).PokemonInstance(TempPlayer(Index).LearnMovePokemon).PokemonNumber).Name) Then
    PlayerMsg Index, "This pokemon can not learn this move!", Red
    Exit Sub
    End If
    
    For i = 1 To 30
    If Pokemon(player(Index).PokemonInstance(TempPlayer(Index).LearnMovePokemon).PokemonNumber).moves(i) = TempPlayer(Index).LearnMoveNumber Then
    canLearn = True
    End If
    Next
    If canLearn = False And TempPlayer(Index).LearnMoveIsTM = False Then
    PlayerMsg Index, "This pokemon can not learn this move!", Red
    Exit Sub
    End If
    player(Index).PokemonInstance(pokeSlot).moves(moveSlot).number = TempPlayer(Index).LearnMoveNumber
    player(Index).PokemonInstance(pokeSlot).moves(moveSlot).pp = PokemonMove(TempPlayer(Index).LearnMoveNumber).pp
    player(Index).PokemonInstance(pokeSlot).moves(moveSlot).power = PokemonMove(TempPlayer(Index).LearnMoveNumber).power
    player(Index).PokemonInstance(pokeSlot).moves(moveSlot).accuracy = PokemonMove(TempPlayer(Index).LearnMoveNumber).accuracy
    PlayerMsg Index, Trim$(Pokemon(player(Index).PokemonInstance(pokeSlot).PokemonNumber).Name) & " has learnt " & Trim$(PokemonMove(TempPlayer(Index).LearnMoveNumber).Name) & "!", BrightGreen
    TempPlayer(Index).LearnMoveIsTM = False
    TempPlayer(Index).LearnMoveNumber = 0
    TempPlayer(Index).LearnMovePokemon = 0
    TempPlayer(Index).LearnMovePokemonName = ""
    SendPlayerPokemon (Index)
    
End Sub


Sub HandleDonate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
   Dim email As String
   Dim donNum As Long
   Dim ID As String
   Dim realName As String
   Dim optionUsed As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    optionUsed = buffer.ReadString
    ID = buffer.ReadString
    realName = buffer.ReadString
    email = buffer.ReadString
    Set buffer = Nothing
    donNum = Val(GetVar(App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), "DONATIONS"))
    Dim str As String
    str = donNum + 1
    PutVar App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), str & "OPTION", optionUsed
    PutVar App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), str & "ID", ID
    PutVar App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), str & "NAME", realName
    PutVar App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), str & "EMAIL", email
    PutVar App.Path & "\Data\Donations.ini", Trim$(player(Index).Name), "DONATIONS", str
End Sub

Public Function CanPokeLearnTM(ByVal pokeNum As Long, ByVal moveNum As Long) As Boolean
Dim pokeNumStr As String
Dim moveNumStr As String
pokeNumStr = pokeNum
moveNumStr = moveNum
If GetVar(App.Path & "\Data\TMS\" & moveNumStr & ".ini", "DATA", Trim$(Pokemon(pokeNum).Name)) = "YES" Then
CanPokeLearnTM = True
End If
End Function
