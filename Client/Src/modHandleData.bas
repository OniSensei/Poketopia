Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SClassesData) = GetAddress(AddressOf HandleClassesData)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHp) = GetAddress(AddressOf HandlePlayerHp)
    HandleDataSub(SPlayerMp) = GetAddress(AddressOf HandlePlayerMp)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNpcDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNpcEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNpc)
    HandleDataSub(SMapKey) = GetAddress(AddressOf HandleMapKey)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SPokemonEditor) = GetAddress(AddressOf HandlePokemonEditor)
    HandleDataSub(SUpdatePokemon) = GetAddress(AddressOf HandleUpdatePokemon)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SDoorAnimation) = GetAddress(AddressOf HandleDoorAnimation)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    'HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNpcVitals)
    HandleDataSub(SCooldown) = GetAddress(AddressOf HandleCooldown)
    'HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(SNpcBattle) = GetAddress(AddressOf HandleNpcBattle)
    HandleDataSub(SPlayerPokemon) = GetAddress(AddressOf HandlePlayerPokemon)
    HandleDataSub(SBattleUpdate) = GetAddress(AddressOf HandleBattleUpdate)
    HandleDataSub(SBattleMessage) = GetAddress(AddressOf HandleBattleMessage)
    HandleDataSub(SMovesEditor) = GetAddress(AddressOf HandleMovesEditor)
    HandleDataSub(SUpdateMove) = GetAddress(AddressOf HandleUpdateMove)
    HandleDataSub(SSound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SOpenStorage) = GetAddress(AddressOf HandleOpenStorage)
    HandleDataSub(SStorageUpdate) = GetAddress(AddressOf HandleStorageUpdate)
    HandleDataSub(SStorageLoadPoke) = GetAddress(AddressOf HandleStorageLoadPoke)
    HandleDataSub(STrainerCard) = GetAddress(AddressOf HandleTrainerCard)
    HandleDataSub(SOpenBank) = GetAddress(AddressOf HandleOpenBank)
    HandleDataSub(SUpdateBank) = GetAddress(AddressOf HandleUpdateBank)
    HandleDataSub(SOpenRoster) = GetAddress(AddressOf HandleOpenRoster)
    HandleDataSub(SBattleInfo) = GetAddress(AddressOf HandleBattleInfo)
    HandleDataSub(SOpenSwitch) = GetAddress(AddressOf HandleOpenSwitch)
    HandleDataSub(SPCRequest) = GetAddress(AddressOf HandlePCRequest)
    HandleDataSub(SPCScan) = GetAddress(AddressOf HandlePcScan)
    HandleDataSub(SisInBattle) = GetAddress(AddressOf HandleisInbattle)
    HandleDataSub(SIntro) = GetAddress(AddressOf HandleIntro)
    HandleDataSub(SMapMusic) = GetAddress(AddressOf HandleMapMusic)
    HandleDataSub(SVersionCheck) = GetAddress(AddressOf HandleVersionCheck)
    HandleDataSub(SAdminCheck) = GetAddress(AddressOf HandleAdminCheck)
    HandleDataSub(STotalPlayersCheck) = GetAddress(AddressOf HandleTotalPlayersCheck)
    'HandleDataSub(SCustomMap) = GetAddress(AddressOf HandleCustomMap)
    HandleDataSub(SDialogs) = GetAddress(AddressOf HandleDialogg)
    HandleDataSub(SNPCScript) = GetAddress(AddressOf HandleNPCScript)
    HandleDataSub(SSend) = GetAddress(AddressOf HandlePacketData)
     HandleDataSub(STPRemove) = GetAddress(AddressOf HandleTPRemove)
     HandleDataSub(SPVPCommand) = GetAddress(AddressOf HandlePVPCommand)
     HandleDataSub(SCrew) = GetAddress(AddressOf HandleCrewData)
      HandleDataSub(SJournal) = GetAddress(AddressOf HandleJournal)
  
End Sub

Public Sub HandleOpenSwitch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
UpdateBattle
frmBattle.picSwitch.Visible = True
End Sub


Public Sub HandlePCRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
SendPCScanResult
End Sub


Public Sub HandleNPCScript(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
  Dim map As Long
  Dim x As Long
  Dim y As Long
  Dim Name As String
  Dim script As String
  Dim Buffer As clsBuffer
  Set Buffer = New clsBuffer
  Buffer.WriteBytes Data()
  map = Buffer.ReadLong
  x = Buffer.ReadLong
  y = Buffer.ReadLong
  Name = Buffer.ReadString
  script = Buffer.ReadString
  Set Buffer = Nothing
  
  frmEditorMapNPC.Show
  'frmEditorMapNPC.Text1.text = Name
  frmEditorMapNPC.RichTextBox1.text = script
  CurrentNpcX = x
  CurrentNpcY = y

End Sub

Public Sub HandlePcScan(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)





End Sub

Public Sub HandleIntro(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
frmIntro.Show
End Sub


Public Sub HandleStorageLoadPoke(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim slot As Long
  Dim Buffer As clsBuffer
  Set Buffer = New clsBuffer
  Buffer.WriteBytes Data()
  slot = Buffer.ReadLong
  Set Buffer = Nothing

'Load pokemon

If frmStorage.Visible = True Then
storagenum = slot
frmStorage.LoadPokemon (storagenum)
End If
End Sub

Sub HandleBattleInfo(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim ms As Long
Dim pokeName As String
Dim pc As Long
Dim win As Long
Dim EXP As Long
Dim round As Long
Dim i As Long
Dim Buffer As clsBuffer
Set Buffer = New clsBuffer
Buffer.WriteBytes Data
ms = Buffer.ReadLong
pokeName = Buffer.ReadString
pc = Buffer.ReadLong
win = Buffer.ReadLong
EXP = Buffer.ReadLong
Set Buffer = Nothing
'Update gui
If win = 0 Then Exit Sub
frmBattle.lblInfoPoke.Caption = Trim$(pokeName)
frmBattle.lblInfoChance.Caption = "1 of " & ms
frmBattle.lblInfoPC.Caption = pc
frmBattle.lblInfoExp.Caption = EXP
AddBattleText "____________________", White
AddBattleText "Battle Info:", Yellow
AddBattleText "Chance: 1 of " & ms, White
AddBattleText "PokeCoins earned: " & pc, White
AddBattleText "EXP earned: " & EXP, BrightCyan
Select Case win
Case YES
AddBattleText "Victory!", BrightGreen
Case BATTLE_NO
AddText "You lost the battle!", BrightRed
frmBattle.Visible = False
frmBattle.picBattleInfo.Visible = False
frmMainGame.picBattleCommands.Visible = False
frmMainGame.txtBtlLog.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
'frmMainGame.picBattleInfo.Visible = False
inBattle = False
frmMainGame.Enabled = True
frmMainGame.SetFocus
frmBattle.txtBtlLog.text = vbNullString
StopPlay
Exit Sub
Case 3
'AddText "You ran away!", BrightRed
frmBattle.picBattleInfo.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
inBattle = False
Unload frmBattle
frmMainGame.picBattleCommands.Visible = False
frmMainGame.txtBtlLog.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
frmMainGame.Enabled = True
frmMainGame.SetFocus
StopPlay
Exit Sub
Case 4
frmBattle.picBattleInfo.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
inBattle = False
Unload frmBattle
frmMainGame.picBattleCommands.Visible = False
frmMainGame.txtBtlLog.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
frmMainGame.Enabled = True
frmMainGame.SetFocus
StopPlay
Exit Sub
End Select
If Player(MyIndex).inBattle = False Then
'CanMoveNow = False
frmMainGame.btnCloseBattle.Visible = True
For i = 1 To 4
frmMainGame.cmdPokeMove(i).Visible = False
Next
frmMainGame.cmdAutoClose.Visible = False
frmMainGame.cmdBag.Visible = False
frmMainGame.cmdRun.Visible = False
frmMainGame.lblBattleEXP.Visible = False
If AutoCloseBattle = True Then
CanMoveNow = True
frmMainGame.btnCloseBattle.Visible = False
frmMainGame.txtBtlLog.Visible = False
frmMainGame.txtBtlLog.text = vbNullString
frmMainGame.listBag.Visible = False
inBattle = False
frmMainGame.picBattleCommands.Visible = False
frmMainGame.Enabled = True
frmMainGame.SetFocus
frmMainGame.txtBtlLog.text = vbNullString
StopPlay
PlayMapMusic MapMusic
Unload frmBattle
End If
BlockBattle
End If
'UpdateBattle
End Sub

Sub HandleOpenBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
frmMainGame.OpenMenu (MENU_BANK)
End Sub

Sub HandleUpdateBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim x As Long
Dim y As Long
 Dim Buffer As clsBuffer
  Set Buffer = New clsBuffer
  Buffer.WriteBytes Data()
  x = Buffer.ReadLong
  y = Buffer.ReadLong
  frmMainGame.lblCPC.Caption = x
  Player(Index).StoredPC = y
  frmMainGame.lblSPC.Caption = y
  Set Buffer = Nothing
If frmMainGame.picBank.Visible = True Then
frmMainGame.LoadBank
End If


End Sub


Sub HandleData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgType = Buffer.ReadInteger

    If MsgType < 0 Then
        MsgBox "Packet Error.", vbCritical
        DestroyGame
        Set Buffer = Nothing
        Exit Sub
    End If

    If MsgType >= SMSG_COUNT Then
        'MsgBox "Packet Error: MsgType = " & MsgType & ".", vbCritical
        DestroyGame
        Set Buffer = Nothing
        Exit Sub
    End If
    'CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length), 0, 0
    CallWindowProc HandleDataSub(MsgType), 1, Buffer.ReadBytes(Buffer.Length - 2 + 1), 0, 0
    Set Buffer = Nothing
End Sub
Sub HandleOpenStorage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
frmStorage.Show
End Sub
Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'frmSendGetData.Visible = False
    frmMainGame.lblSGInfo.Visible = False
    frmMenu.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picNewChar.Visible = False
    frmMenu.picRegister.Visible = False
    Msg = Buffer.ReadString 'Parse(1)
    Set Buffer = Nothing
    Call MsgBox(Msg, vbOKOnly, GAME_NAME)
End Sub




Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' save options
    Options.SavePass = frmMenu.chkPass.Value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong 'CLng(Parse(1))
    Set Buffer = Nothing
    'frmSendGetData.Visible = True
     frmMainGame.lblSGInfo.Visible = True
    Call SetStatus("Receiving game data...")
    ReceivingTime = GetTickCount + 60000
End Sub




Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString
            .Vital(Vitals.HP) = Buffer.ReadLong
            .Vital(Vitals.MP) = Buffer.ReadLong
            .Vital(Vitals.SP) = Buffer.ReadLong
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next
            
            .Stat(Stats.strength) = Buffer.ReadLong
            .Stat(Stats.endurance) = Buffer.ReadLong
            .Stat(Stats.vitality) = Buffer.ReadLong
            .Stat(Stats.intelligence) = Buffer.ReadLong
            .Stat(Stats.willpower) = Buffer.ReadLong
            .Stat(Stats.spirit) = Buffer.ReadLong
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMainGame.picNewChar.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMainGame.Picture1.Visible = False
    'frmSendGetData.Visible = False
     frmMainGame.lblSGInfo.Visible = False
    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).Name)
    Next

    frmMenu.cmbClass.ListIndex = 0
    n = frmMenu.cmbClass.ListIndex + 1
    
    newCharSprite = 1
    NewCharacterBltSprite (newCharSprite)
End Sub

Sub HandleClassesData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(1 To Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))
            .Vital(Vitals.HP) = Buffer.ReadLong 'CLng(Parse(n + 1))
            .Vital(Vitals.MP) = Buffer.ReadLong 'CLng(Parse(n + 2))
            .Vital(Vitals.SP) = Buffer.ReadLong 'CLng(Parse(n + 3))
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next
                            
            .Stat(Stats.strength) = Buffer.ReadLong 'CLng(Parse(n + 4))
            .Stat(Stats.endurance) = Buffer.ReadLong
            .Stat(Stats.vitality) = Buffer.ReadLong
            .Stat(Stats.intelligence) = Buffer.ReadLong
            .Stat(Stats.willpower) = Buffer.ReadLong
            .Stat(Stats.spirit) = Buffer.ReadLong
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
If InGame = False Then
    InGame = True
    Call GameInit
    Call GameLoop
    ReceivingTime = 0
    End If
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(Index, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(Index, i, Buffer.ReadLong)
        n = n + 2
    Next
    
    If Index = MyIndex Then
        ' changes to inventory, need to clear any drop menu
        'frmMainGame.picCurrency.Visible = False
        'frmMainGame.txtCurrency.text = vbNullString
        'tmpCurrencyItem = 0
        'CurrencyMenu = 0 ' clear

        frmBag.LoadInv
        frmMainGame.BagLoadInv
        If frmPokemons.Visible = True Then frmPokemons.LoadInv
        If frmMainGame.picPokemons.Visible = True Then frmMainGame.RosterLoadInv
    End If

    Set Buffer = Nothing
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(Index, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(Index, n, Buffer.ReadLong) 'CLng(Parse(3)))

    If Index = MyIndex Then
        
    
        frmBag.LoadInv
        frmMainGame.BagLoadInv
        If frmPokemons.Visible = True Then frmPokemons.LoadInv
        If frmMainGame.picPokemons.Visible = True Then frmMainGame.RosterLoadInv
    End If

    Set Buffer = Nothing
    If GetPlayerInvItemNum(Index, n) = 1 Then
    
    If frmMainGame.picBank.Visible = True Then
    frmMainGame.LoadBank
    End If
    End If
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'Call SetPlayerEquipment(Index, Buffer.ReadLong, Armor)
    'Call SetPlayerEquipment(Index, Buffer.ReadLong, Weapon)
    'Call SetPlayerEquipment(Index, Buffer.ReadLong, Helmet)
    'Call SetPlayerEquipment(Index, Buffer.ReadLong, Shield)
    
    If Index = MyIndex Then
        ' changes to inventory, need to clear any drop menu
        
    
        BltInventory
        frmBag.LoadInv
        frmMainGame.BagLoadInv
        frmPokemons.LoadInv
         frmMainGame.RosterLoadInv
        BltEquipment
    End If

    Set Buffer = Nothing
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim playernum As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    playernum = Buffer.ReadLong
    'Call SetPlayerEquipment(playernum, Buffer.ReadLong, Armor)
    'Call SetPlayerEquipment(playernum, Buffer.ReadLong, Weapon)
    'Call SetPlayerEquipment(playernum, Buffer.ReadLong, Helmet)
    'Call SetPlayerEquipment(playernum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerHp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(Index).MaxHp = Buffer.ReadLong
    Call SetPlayerVital(Index, Vitals.HP, Buffer.ReadLong)

   
End Sub


Private Sub HandlePlayerMp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(Index).MaxMP = Buffer.ReadLong
    Call SetPlayerVital(Index, Vitals.MP, Buffer.ReadLong)

End Sub

Private Sub HandlePlayerSp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(Index).MaxSP = Buffer.ReadLong
    Call SetPlayerVital(Index, Vitals.SP, Buffer.ReadLong)
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim TNL As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SetPlayerExp Index, Buffer.ReadLong
    TNL = Buffer.ReadLong
    If Index = MyIndex Then
        
    End If
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, x As Long, a As Long, b As Long, nv As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Player(i).mood = Buffer.ReadLong
    nv = Buffer.ReadLong
    For b = 1 To 6
    Player(i).Pokes(b) = Buffer.ReadLong
    Next
    For x = 1 To Stats.stat_count - 1
        SetPlayerStat i, x, Buffer.ReadLong
    Next

For a = 1 To MAX_GYMS
 Player(i).Badge(a) = Buffer.ReadLong
Next
Player(i).Equipment(Equipment.Armor) = Buffer.ReadLong
Player(i).Equipment(Equipment.Helmet) = Buffer.ReadLong
Player(i).Equipment(Equipment.Shield) = Buffer.ReadLong
Player(i).Equipment(Equipment.Weapon) = Buffer.ReadLong
Player(i).Equipment(Equipment.mask) = Buffer.ReadLong
Player(i).Equipment(Equipment.Outfit) = Buffer.ReadLong
Player(i).HasBike = Buffer.ReadLong
    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        
        ' Set the character windows
    
        
       
   
    End If
    If Player(MyIndex).Access < 1 Then
    frmMainGame.btnAdminPanel.Caption = "Acces needed"
    frmMainGame.btnAdminPanel.Visible = False
    Else
    frmMainGame.btnAdminPanel.Caption = "Admin Panel"
    frmMainGame.btnAdminPanel.Visible = True
    End If
    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    If nv = YES Then
    Player(i).notVisible = True
    Else
    Player(i).notVisible = False
    End If
    If frmMainGame.picProfile.Visible = True Then frmMainGame.loadClothes
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Call SetPlayerX(i, x)
    Call SetPlayerY(i, y)
    Call SetPlayerDir(i, dir)
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).YOffset = PIC_Y
        Case DIR_DOWN
            Player(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).XOffset = PIC_X
        Case DIR_RIGHT
            Player(i).XOffset = PIC_X * -1
    End Select
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNpcNum As Long
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Movement = Buffer.ReadLong

    With MapNpc(MapNpcNum)
        .x = x
        .y = y
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
        End Select

    End With

End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerDir(i, dir)

    With Player(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim dir As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    dir = Buffer.ReadLong

    With MapNpc(i)
        .dir = dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim dir As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    dir = Buffer.ReadLong
    Call SetPlayerX(Index, x)
    Call SetPlayerY(Index, y)
    Call SetPlayerDir(Index, dir)
    ' Make sure they aren't walking
    Player(Index).Moving = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim NeedMap As Byte
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearTempTile
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    ' Get map num
    x = Buffer.ReadLong
    ' Get revision
    y = Buffer.ReadLong

    If FileExist(MAP_PATH & "map" & x & MAP_EXT, False) Then
        Call LoadMap(x)
        ' Check to see if the revisions match
        NeedMap = 1

        If map.Revision = y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong TCP_CODE
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim MapNum As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    Buffer.DecompressBuffer
    
    n = 1
    
    MapNum = Buffer.ReadLong
    map.Name = Buffer.ReadString
    map.Revision = Buffer.ReadLong 'CLng(Parse(n + 2))
    map.Moral = Buffer.ReadLong 'CByte(Parse(n + 3))
    map.tileset = Buffer.ReadLong 'CInt(Parse(n + 4))
    map.Up = Buffer.ReadLong 'CInt(Parse(n + 5))
    map.Down = Buffer.ReadLong 'CInt(Parse(n + 6))
    map.Left = Buffer.ReadLong 'CInt(Parse(n + 7))
    map.Right = Buffer.ReadLong 'CInt(Parse(n + 8))
    map.music = Buffer.ReadLong
    map.BootMap = Buffer.ReadLong
    map.BootX = Buffer.ReadLong
    map.BootY = Buffer.ReadLong
    map.MaxX = Buffer.ReadLong
    map.MaxY = Buffer.ReadLong
    
    ReDim map.Tile(0 To map.MaxX, 0 To map.MaxY)
    n = n + 16

    For x = 0 To map.MaxX
        For y = 0 To map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                map.Tile(x, y).Layer(i).x = Buffer.ReadByte
                map.Tile(x, y).Layer(i).y = Buffer.ReadByte
                map.Tile(x, y).Layer(i).tileset = Buffer.ReadByte
            Next
            map.Tile(x, y).Type = Buffer.ReadLong 'CByte(Parse(n + 6))
            map.Tile(x, y).data1 = Buffer.ReadLong 'CInt(Parse(n + 7))
            map.Tile(x, y).data2 = Buffer.ReadLong 'CInt(Parse(n + 8))
            map.Tile(x, y).data3 = Buffer.ReadLong 'CInt(Parse(n + 9))
            n = n + 10
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        map.NPC(x) = Buffer.ReadLong 'CByte(Parse(n))
        n = n + 1
    Next

    For x = 1 To MAX_MAP_POKEMONS
        map.Pokemon(x).PokemonNumber = Buffer.ReadLong
        map.Pokemon(x).LevelFrom = Buffer.ReadLong
        map.Pokemon(x).LevelTo = Buffer.ReadLong
        map.Pokemon(x).Custom = Buffer.ReadLong
        map.Pokemon(x).ATK = Buffer.ReadLong
        map.Pokemon(x).DEF = Buffer.ReadLong
        map.Pokemon(x).SPATK = Buffer.ReadLong
        map.Pokemon(x).SPDEF = Buffer.ReadLong
        map.Pokemon(x).SPD = Buffer.ReadLong
        map.Pokemon(x).HP = Buffer.ReadLong
        map.Pokemon(x).Chance = Buffer.ReadLong
    Next
    ClearTempTile
    
    Set Buffer = Nothing
    
    ' Save the map
    Call SaveMap(MapNum) 'CLng(Parse(1)))
    
    
    'Load maps next to this one
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
    

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmMapProperties.Visible Then
            Unload frmMapProperties
        End If
    End If

End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS

        With MapItem(i)
            .num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
        End With

    Next

End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS

        With MapNpc(i)
            .num = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            .dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With

    Next

End Sub

Private Sub HandleMapDone()
    Dim i As Long
    Dim MusicFile As String
    
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    
    ' load tilesets we need
    LoadTilesets
            
    MusicFile = Trim$(CStr(map.music)) & ".mid"
    Call UpdateDrawMapName

    GettingMap = False
    CanMoveNow = True
End Sub

Private Sub HandleBroadcastMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
    If frmChat.txtMyChat.text = "" Then
    frmChat.txtMyChat.Visible = True
    frmChat.txtMyChat.Enabled = True
    frmChat.txtChat.Visible = True
    frmChat.txtChat.Enabled = True
    isChatVisible = True
    End If
End Sub

Private Sub HandleTrainerCard(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim PPic As String
    Dim i As Long
    Dim a As Long
    Dim b As Long
    Dim gym As Long
    Dim gymheader As String
    Dim pokemonnum(1 To 6) As Integer
    Dim pokelvl(1 To 6) As Integer
    Dim isShiny(1 To 6) As Integer
    Dim badges(1 To 6) As Integer
    Dim plvl As Long
    Dim RankPoints As Long
    Dim Crew As String
    Dim crewPic As String
    Dim doCrew As Long
    plvl = 0
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Name = Buffer.ReadString
    For i = 1 To 6
    pokemonnum(i) = Buffer.ReadLong
    pokelvl(i) = Buffer.ReadLong
    isShiny(i) = Buffer.ReadLong
    badges(i) = Buffer.ReadLong
    Next
    PPic = Buffer.ReadString
    RankPoints = Buffer.ReadLong
    Crew = Buffer.ReadString
    crewPic = Buffer.ReadString
    doCrew = Buffer.ReadLong
    'Load GUI
    frmMainGame.OpenMenu (MENU_TRAINERCARD)
    
    If GetPlayerAccess(FindPlayer(Name)) > 0 Then
frmMainGame.GameMaster.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\GM.png")
Call frmMainGame.GameMaster.SetFixedSizeAspect(frmMainGame.GameMaster.Width / 15, frmMainGame.GameMaster.Height / 15, True)
Else
frmMainGame.GameMaster.Picture = Nothing
End If
    For b = 1 To MAX_GYMS
    If Player(FindPlayer(Name)).Bedages(b) = GYM_DEFEATED Then
    If b > gym Then
    gym = b
    End If
    End If
    Next
    For i = 1 To 6
    If badges(i) = YES Then

    frmMainGame.imgBadge(i).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\badges\" & i & ".png")
    Else
    frmMainGame.imgBadge(i).Picture = Nothing
    End If
    Next

    
    
    
    frmMainGame.lblName.Caption = Trim$(Name)
    If FindPlayer(Trim$(Name)) = MyIndex Then

       
        frmMainGame.lvButtons_H11.Visible = False
    frmMainGame.lvButtons_H1.Visible = False
        
       Else
     
         frmMainGame.lvButtons_H11.Visible = True
       frmMainGame.lvButtons_H1.Visible = True
    End If
    
    For a = 1 To 6
    If pokemonnum(a) <= 0 Then
    Set frmMainGame.imgCharPokemon(a).Picture = Nothing
        frmMainGame.lblCharPkmnLvl(a).Caption = "Lvl:0"
    Else
      If isShiny(a) = YES Then
       frmMainGame.imgCharPokemon(a).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & pokemonnum(a) & ".gif")
      Else
       frmMainGame.imgCharPokemon(a).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & pokemonnum(a) & ".gif")
      End If
       
        frmMainGame.lblCharPkmnLvl(a).Caption = "Lvl:" & pokelvl(a)
        plvl = plvl + pokelvl(a)
    End If
    Next
    If PPic <> "" Then
    frmMainGame.imgProfilePic.Picture = LoadPictureGDIplus(PPic)
    
    Else
    frmMainGame.imgProfilePic.Picture = LoadPictureGDIplus("http://orig06.deviantart.net/4b80/f/2012/276/2/1/nate_icon_by_pheonixmaster1-d5go0io.png")
    End If
    
     frmMainGame.lblCharPowerLvl.Caption = "Power Lvl:" & plvl
     
     If RankPoints < 5 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-2.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Bronze 3"
     End If
     If RankPoints >= 5 And RankPoints < 15 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-1.png")
      frmMainGame.lblRankPoints.Caption = RankPoints & " - Bronze 2"
     End If
      If RankPoints >= 15 And RankPoints < 25 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-0.png")
      frmMainGame.lblRankPoints.Caption = RankPoints & " - Bronze 1"
     End If
     If RankPoints >= 25 And RankPoints < 35 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-2.png")
      frmMainGame.lblRankPoints.Caption = RankPoints & " - Silver 3"
     End If
      If RankPoints >= 35 And RankPoints < 45 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-1.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Silver 2"
     End If
      If RankPoints >= 45 And RankPoints < 55 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-0.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Silver 1"
     End If
      If RankPoints >= 55 And RankPoints < 65 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-2.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Gold 3"
     End If
     If RankPoints >= 65 And RankPoints < 75 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-1.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Gold 2"
     End If
     If RankPoints >= 75 And RankPoints < 90 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-0.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Gold 1"
     End If
     If RankPoints >= 90 And RankPoints < 105 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-2.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Platinum 3"
     End If
     If RankPoints >= 105 And RankPoints < 120 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-1.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Platinum 2"
     End If
     If RankPoints >= 120 And RankPoints < 150 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-0.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Platinum 1"
     End If
     If RankPoints >= 150 And RankPoints < 175 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-2.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Diamond 2"
     End If
      If RankPoints >= 175 And RankPoints < 200 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-1.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Diamond 1"
     End If
      If RankPoints >= 200 Then
     frmMainGame.imgRank = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-0.png")
     frmMainGame.lblRankPoints.Caption = RankPoints & " - Champion"
     End If
     frmMainGame.cmdClanInvite.Visible = False
     frmMainGame.lblPlayerCrew.Caption = "Clan - " & Crew
     If crewPic = "" Then
     frmMainGame.imgPlayerClan.Picture = Nothing
     Else
     frmMainGame.imgPlayerClan.Picture = LoadPictureGDIplus(crewPic)
     End If
     If Crew = "None" Then
     If doCrew = YES Then
     frmMainGame.cmdClanInvite.Visible = True
     Else
     frmMainGame.cmdClanInvite.Visible = False
     End If
     End If
     
     frmMainGame.picTrainerCard.Visible = True
     
End Sub


Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim Sound As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Sound = Buffer.ReadString
    Call PlaySound(Sound)
    
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
End Sub

Private Sub HandleMapMusic(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim Buffer As clsBuffer
    Dim music As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    music = Buffer.ReadString
    Set Buffer = Nothing
    MapMusic = Trim$(music)
End Sub

Private Sub HandleStarter(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Starter As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
   Starter = Buffer.ReadLong
    Set Buffer = Nothing
   WaitingStarter = Starter
   frmChoose.Show
   frmChoose.imgPoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & Starter & ".gif")
   frmChoose.imgPoke.Animate (lvicAniCmdStart)
   frmChoose.imgType1.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(Starter).Type & ".bmp")
   frmChoose.imgType2.Picture = LoadPicture(App.Path & "\Data Files\graphics\types\" & Pokemon(Starter).Type2 & ".bmp")
End Sub


Private Sub HandleDialogg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    Dim image As Long
    Dim dialogtxt As String
    Dim isTrigger As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    dialogtxt = Buffer.ReadString
    image = Buffer.ReadLong
    isTrigger = Buffer.ReadLong
    Set Buffer = Nothing
    
    If CurrentDialog = 0 And Dialogs = 0 Then
    
    CanMoveNow = False
    CurrentDialog = 1
    Dialog(1) = dialogtxt
    If isTrigger = YES Then
    IsDialogTrigger(1) = True
    Else
    IsDialogTrigger(1) = False
    End If
    Dialogs = 1
    frmMainGame.picDialog.Visible = True
    If image > 0 Then
    frmMainGame.picDialog.Left = 15
    DialogImage(1) = image
    If FileExist("Data Files\pictures\" & image & ".png") Then
   
    Else
    frmMainGame.picDialog.Left = 120
    frmMainGame.imgDialogPic = Nothing
    End If
    Else
    frmMainGame.imgDialogPic = Nothing
    frmMainGame.picDialog.Left = 120
    End If
    frmMainGame.DisplayDialogText Dialog(1)
    'frmMainGame.txtDialog.Caption = Dialog(1)
    
    Else
    Dialog(Dialogs + 1) = dialogtxt
    DialogImage(Dialogs + 1) = image
    If isTrigger = YES Then
    IsDialogTrigger(Dialogs + 1) = True
    Else
    IsDialogTrigger(Dialogs + 1) = False
    End If
    Dialogs = Dialogs + 1
    End If
End Sub


Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Msg As String
    Dim color As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Call AddText(Msg, color)
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
    End With

End Sub

Private Sub HandleItemEditor()
    Dim i As Long
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    ' changes to inventory, need to clear any drop menu
   
    tmpDropItem = 0
    'BltInventory
    frmBag.LoadInv
    frmMainGame.BagLoadInv
    frmPokemons.LoadInv
    frmMainGame.RosterLoadInv
   
    
End Sub


Private Sub HandleisInbattle(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim n As Long
    Dim indx As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    indx = Buffer.ReadLong
    n = Buffer.ReadLong
    Set Buffer = Nothing
    If n = YES Then
    Player(indx).inBattle = YES
    Else
    Player(indx).inBattle = NO
    End If
    
   
    
End Sub


Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = Buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set Buffer = Nothing
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .num = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .dir = Buffer.ReadLong
        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

End Sub

Private Sub HandleNpcDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
End Sub

Private Sub HandleNpcEditor()
    Dim i As Long
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With

End Sub

Private Sub HandleUpdateNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim n As Long
    Dim Buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ' Update the Npc
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = Buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    Set Buffer = Nothing
End Sub


Private Sub HandleResourceEditor()
    Dim i As Long

    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ResourceNum = Buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = Buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePokemonEditor()
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
Dim i As Long

    With frmEditor_Pokemon
        Editor = EDITOR_POKEMON
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_POKEMONS
            .lstIndex.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        PokemonEditorInit
    End With

End Sub

Private Sub HandleMovesEditor()
    Dim i As Long
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
    With frmEditor_Moves
        Editor = EDITOR_POKEMON
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_MOVES
            .lstIndex.AddItem i & ": " & Trim$(PokemonMove(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        MovesEditorInit
    End With

End Sub

Private Sub HandleUpdatePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim pokemonnum As Long
    Dim Buffer As clsBuffer
    Dim PokemonSize As Long
    Dim PokemonData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    pokemonnum = Buffer.ReadLong
    
    PokemonSize = LenB(Pokemon(pokemonnum))
    ReDim PokemonData(PokemonSize - 1)
    PokemonData = Buffer.ReadBytes(PokemonSize)
    CopyMemory ByVal VarPtr(Pokemon(pokemonnum)), ByVal VarPtr(PokemonData(0)), PokemonSize
    
    Set Buffer = Nothing
End Sub
Private Sub HandleUpdateMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
 
    Dim movenum As Long
    Dim Buffer As clsBuffer
    Dim MoveSize As Long
    Dim MoveData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    movenum = Buffer.ReadLong
    
    MoveSize = LenB(PokemonMove(movenum))
    ReDim MoveData(MoveSize - 1)
    MoveData = Buffer.ReadBytes(MoveSize)
    CopyMemory ByVal VarPtr(PokemonMove(movenum)), ByVal VarPtr(MoveData(0)), MoveSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleMapKey(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    n = Buffer.ReadLong
    TempTile(x, y).DoorOpen = n
End Sub

Private Sub HandleEditMap()
    Call MapEditorInit
End Sub

Private Sub HandleShopEditor()
    Dim i As Long
If GetPlayerAccess(MyIndex) < 2 Then Exit Sub
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim shopnum As Long
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopnum)), ByVal VarPtr(ShopData(0)), ShopSize
    
    Set Buffer = Nothing
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    spellnum = Buffer.ReadLong
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = Buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set Buffer = Nothing
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next
    
    BltPlayerSpells
    Set Buffer = Nothing
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resource_Index = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = Buffer.ReadByte
            MapResource(i).x = Buffer.ReadLong
            MapResource(i).y = Buffer.ReadLong
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart
End Sub

Private Sub HandleDoorAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    With TempTile(x, y)
        .DoorFrame = 1
        .DoorAnimate = 1 ' 0 = nothing| 1 = opening | 2 = closing
        .DoorTimer = GetTickCount
    End With
    Set Buffer = Nothing
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, message As String, color As Long, tmpType As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg message, color, tmpType, x, y
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, Sprite As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, 3)
    
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .x = x
        .y = y
        .Sprite = Sprite
        .Timer = GetTickCount
    End With
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With

    Set Buffer = Nothing
End Sub

Private Sub HandleMapNpcVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadByte
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
End Sub

Private Sub HandleCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim slot As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    slot = Buffer.ReadLong
    SpellCD(slot) = GetTickCount
    
    BltPlayerSpells
    
    Set Buffer = Nothing
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SpellBuffer = 0
    SpellBufferTimer = 0
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim Access As Long
    Dim Name As String
    Dim message As String
    Dim Colour As Long
    Dim Header As String
    Dim PK As Long
    Dim saycolour As Long
    Dim acs As String
    Dim pcslot As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    Header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                
                acs = "[Player]"
                Colour = QBColor(White)
                
            
            Case 1
                Colour = QBColor(BrightCyan)
                acs = "[MOD]"
            Case 2
                Colour = QBColor(BrightCyan)
                acs = "[MOD]"
            Case 3
                Colour = QBColor(BrightCyan)
                acs = "[MOD+]"
            Case 4
                Colour = QBColor(BrightRed)
                acs = "[Admin]"
                If Name = "Goran" Then
                acs = "[Admin][Dev.]"
                End If
        End Select
    Else
        Colour = QBColor(BrightRed)
    End If
    For i = 1 To MAX_INV
    If GetPlayerInvItemNum(FindPlayer(Name), i) = 1 Then
    pcslot = i
    Exit For
    End If
    Next
    
    Select Case Trim$(Header)
    Case "[Global]"
   
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text)
    
    If GetPlayerInvItemValue(FindPlayer(Name), pcslot) >= 1000000 Then
    frmChat.txtChat.SelText = vbNewLine
     If Access >= 1 Then
    'Call AddPicture(frmMainGame.txtChat, "gm.bmp")
    End If
    'frmMainGame.txtChat.SelColor = colour
    'frmMainGame.txtChat.SelText = "[RICH]" & acs & Name & ": " 'vbNewLine & before rich
    Else
    
    frmChat.txtChat.SelText = vbNewLine
     If Access >= 1 Then
    'Call AddPicture(frmMainGame.txtChat, "gm.bmp")
    End If
    frmChat.txtChat.SelColor = Colour
    frmChat.txtChat.SelText = acs & Name & ": " ' vbNewLine & before acs
    End If
    
    frmChat.txtChat.SelColor = saycolour
    frmChat.txtChat.SelText = message
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text) - 1
    ReOrderChat acs & Name & ": " & message, Colour
    
    Case "[Map]"
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text)
    frmChat.txtChat.SelColor = QBColor(White)
    frmChat.txtChat.SelText = vbNewLine & Name & ": "
    frmChat.txtChat.SelColor = QBColor(White)
    frmChat.txtChat.SelText = message
    frmChat.txtChat.SelStart = Len(frmChat.txtChat.text) - 1
    ReOrderChat Name & ": " & message, Colour
    End Select
    
    
        
    Set Buffer = Nothing
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim shopnum As Long
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopnum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop shopnum
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ShopAction = 0
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    StunDuration = Buffer.ReadLong
    
    Set Buffer = Nothing
End Sub

Private Sub HandleVersionCheck(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim a As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    a = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If a = VersionCode Then
    frmMenu.lblUpdate.Visible = True
    Else
    Call MsgBox("Please update your Client!", , "")
    Call DestroyGame
    End If
End Sub

Private Sub HandleTotalPlayersCheck(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim a As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    a = Buffer.ReadLong
    
    Set Buffer = Nothing
    frmMenu.lblPlayers.Caption = a & "/" & MAX_PLAYERS & " players"
End Sub


Private Sub HandleAdminCheck(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim a As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    a = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If a = YES Then
    AdminOnly = True
    Else
    AdminOnly = False
    If InGame = True Then
    Dim Bufr As clsBuffer
    Dim i As Long
    
    isLogging = True
    InGame = False
    
    Set Bufr = New clsBuffer
    Bufr.WriteLong CQuit
    SendData Bufr.ToArray()
    Set Bufr = Nothing
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
     DragInvSlotNum = 0
     InvX = 0
     InvY = 0
     EqX = 0
     EqY = 0
     SpellX = 0
     SpellY = 0
     LastItemDesc = 0
     MyIndex = 0
     InventoryItemSelected = 0
     SpellBuffer = 0
     SpellBufferTimer = 0
     tmpDropItem = 0
    
    frmChat.txtChat.text = vbNullString
    End If
    End If
End Sub




Private Sub HandleNpcBattle(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim isnpc As Long
    Dim NPM As Long
    Dim NM As String
    Dim cB As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'Enemy pokemon
    enemyPokemon.PokemonNumber = Buffer.ReadLong
    enemyPokemon.Level = Buffer.ReadLong
    enemyPokemon.HP = Buffer.ReadLong
    enemyPokemon.MaxHp = Buffer.ReadLong
    enemyPokemon.isShiny = Buffer.ReadLong
    'My pokemon
    BattlePokemon = Buffer.ReadLong
    BattleRound = Buffer.ReadLong
    BattleType = 1 ' npc battle
    isnpc = Buffer.ReadLong
    NPM = Buffer.ReadLong
    NM = Buffer.ReadString
    cB = Buffer.ReadString
    If frmMainGame.menuLeft Then frmMainGame.tmrmenu.Enabled = True
    If isnpc = YES Then
    inBattle = True
    LoadBattleGDI
    
    frmMainGame.picBattleCommands.Visible = True
    frmMainGame.txtBtlLog.Visible = True
    For i = 1 To 4
frmMainGame.cmdPokeMove(i).Visible = True
Next
frmMainGame.cmdAutoClose.Visible = True
frmMainGame.cmdBag.Visible = True
frmMainGame.cmdRun.Visible = True
frmMainGame.lblBattleEXP.Visible = True
frmMainGame.btnCloseBattle.Visible = False
    
    
    If Trim$(cB) <> "" Then
    frmBattle.Picture1.Picture = LoadPicture(App.Path & "\Data Files\graphics\battle\" & cB)
    End If
    'frmMainGame.Enabled = False
    AddBattleText "Opponent sent out " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name) & " Lvl." & enemyPokemon.Level & "!", BrightRed
    'TextAdd frmBattle.txtBattleLog, "A wild " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name) & " has appeared!", True
    If NPM = YES Then
    GoranPlay App.Path & "\Data Files\music\" & NM
    End If
    UpdateBattle
    unBlockBattle
    Else
    inBattle = True
    LoadBattleGDI
    frmMainGame.picBattleCommands.Visible = True
    frmMainGame.txtBtlLog.Visible = True
    frmMainGame.picBattleCommands.Visible = True
    frmMainGame.txtBtlLog.Visible = True
    For i = 1 To 4
frmMainGame.cmdPokeMove(i).Visible = True
Next
frmMainGame.cmdAutoClose.Visible = True
frmMainGame.cmdBag.Visible = True
frmMainGame.cmdRun.Visible = True
frmMainGame.lblBattleEXP.Visible = True
frmMainGame.btnCloseBattle.Visible = False
    'frmMainGame.Enabled = False
    AddBattleText "A wild " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name) & " Lvl." & enemyPokemon.Level & " appeared!", BrightRed
    'TextAdd frmBattle.txtBattleLog, "A wild " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name) & " has appeared!", True
    GoranPlay App.Path & "\Data Files\music\Battle2.mp3"
    UpdateBattle
    unBlockBattle
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandlePlayerPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim i As Long
    Dim x As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Update the pokemon
    For i = 1 To 6
        PokemonInstance(i).PokemonNumber = Buffer.ReadLong
        PokemonInstance(i).Level = Buffer.ReadLong
        PokemonInstance(i).HP = Buffer.ReadLong
        PokemonInstance(i).MaxHp = Buffer.ReadLong
        PokemonInstance(i).pp = Buffer.ReadLong
        PokemonInstance(i).EXP = Buffer.ReadLong
        PokemonInstance(i).TP = Buffer.ReadLong
        PokemonInstance(i).ATK = Buffer.ReadLong
        PokemonInstance(i).DEF = Buffer.ReadLong
        PokemonInstance(i).SPATK = Buffer.ReadLong
        PokemonInstance(i).SPDEF = Buffer.ReadLong
        PokemonInstance(i).SPD = Buffer.ReadLong
        PokemonInstance(i).isShiny = Buffer.ReadLong
        PokemonInstance(i).HoldingItem = Buffer.ReadLong
        For x = 1 To 4
        PokemonInstance(i).moves(x).number = Buffer.ReadLong
        PokemonInstance(i).moves(x).pp = Buffer.ReadLong
        Next
        PokemonInstance(i).nature = Buffer.ReadLong
        PokemonInstance(i).expNeeded = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    ' update GUI
   
    
    'Check if Pokemons form is opened if its opened update it
    If frmPokemons.Visible = True Then
    frmPokemons.loadImages
    frmPokemons.LoadPokemon (selectedpoke)
    End If
    
    If frmMainGame.picPokemons.Visible = True Then
    frmMainGame.RosterloadImages
    frmMainGame.RosterLoadPokemon (selectedpoke)
    End If
    
    For i = 1 To 6
    If PokemonInstance(i).PokemonNumber > 0 Then
    If PokemonInstance(i).HP > 0 Then
    frmMainGame.imgSwitch(i).Picture = Nothing
    frmMainGame.imgSwitch(i).Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokeicons\" & PokemonInstance(i).PokemonNumber & ".png")
    frmMainGame.imgSwitch(i).AnimateOnLoad = True
    Else
    frmMainGame.imgSwitch(i).Picture = Nothing
    End If
    Else
    frmMainGame.imgSwitch(i).Picture = Nothing
    End If
    Next
    
    'Check if battle is on
    If frmBattle.Visible = True Then UpdateBattle
    
End Sub

Private Sub HandleOpenRoster(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
selectedpoke = 1
frmMainGame.RosterLoadPokemon (selectedpoke)
frmMainGame.OpenMenu (MENU_ROSTER)
End Sub

Private Sub HandleStorageUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    ' Update the storage pokemon
    For i = 1 To 250
        StorageInstance(i).PokemonNumber = Buffer.ReadLong
        StorageInstance(i).Level = Buffer.ReadLong
        StorageInstance(i).nature = Buffer.ReadLong
        StorageInstance(i).ATK = Buffer.ReadLong
        StorageInstance(i).DEF = Buffer.ReadLong
        StorageInstance(i).SPD = Buffer.ReadLong
        StorageInstance(i).SPATK = Buffer.ReadLong
        StorageInstance(i).SPDEF = Buffer.ReadLong
        StorageInstance(i).MaxHp = Buffer.ReadLong
        StorageInstance(i).isShiny = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    If frmStorage.Visible = True Then
    storagenum = 1
    frmStorage.LoadPokemon (storagenum)
    initStorage
    End If
    
End Sub

Private Sub HandleBattleUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
Dim i As Long
Dim round As Long
Dim uNblock As Long
Dim oldPOke As Long
Dim oldpokelvl As Long
Dim PvPUB As Long
oldPOke = enemyPokemon.PokemonNumber
oldpokelvl = enemyPokemon.Level
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    'Enemy pokemon
    enemyPokemon.PokemonNumber = Buffer.ReadLong
    enemyPokemon.HP = Buffer.ReadLong
    enemyPokemon.MaxHp = Buffer.ReadLong
    enemyPokemon.Level = Buffer.ReadLong
    enemyPokemon.isShiny = Buffer.ReadLong
    uNblock = Buffer.ReadLong
    PvPUB = Buffer.ReadLong
    'My pokemon
    BattlePokemon = Buffer.ReadLong
    BattleRound = Buffer.ReadLong
    If enemyPokemon.PokemonNumber <= 0 Then
        ' no more battle, exit out
        StopPlay
        BattleType = 0 ' none
       unBlockBattle
       Set Buffer = Nothing
        Exit Sub
    End If
    UpdateBattle
    If InPVP Then
    If PvPUB = YES Then
    unBlockBattle
    End If
    Else
    unBlockBattle
    End If
    If enemyPokemon.Level <> oldpokelvl Or oldPOke <> enemyPokemon.PokemonNumber Then
    AddBattleText "Opponent sent out " & Trim$(Pokemon(enemyPokemon.PokemonNumber).Name) & " Lvl." & enemyPokemon.Level & "!", BrightRed
    End If
    
    Set Buffer = Nothing
End Sub

Private Sub HandleBattleMessage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim color As Long
    Dim Msg As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    color = Buffer.ReadLong
    Set Buffer = Nothing
    AddBattleText Msg, color
    
End Sub



'////////////////////NEW PACKET PROTOCOL////////////////////////////////

Private Sub HandlePacketData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

 Dim i As Long
 Dim Strs As Long
 Dim Longs As Long
 Dim Buffer As clsBuffer
 Dim packetType As String
 Dim stringTrash As String
 Dim longTrash As Long
 Set Buffer = New clsBuffer
 Buffer.WriteBytes Data()
 packetType = Buffer.ReadString
 Strs = Buffer.ReadLong
 Longs = Buffer.ReadLong
 '----------------------
 If Strs > 0 Then
 Dim SendedStrings(1 To 50) As String
 For i = 1 To Strs
 SendedStrings(i) = Buffer.ReadString
 Next
 Else
 stringTrash = Buffer.ReadString
 stringTrash = ""
 End If
 '----------------------
 If Longs > 0 Then
 Dim SendedLongs(1 To 50) As Long
 For i = 1 To Longs
 SendedLongs(i) = Buffer.ReadLong
 Next
 Else
 longTrash = Buffer.ReadLong
 longTrash = 0
 End If
 '----------------------

 
 Select Case packetType
   Case "NPCSCRIPT"
   frmEditorMapNPC.RichTextBox1.text = SendedStrings(1)
   Case "GOLDMSG"
   Case "RSHOP"
   frmMainGame.ShoploadMyItems
   Case "STARTTRADE"
   frmMainGame.txtItem.text = ""
   frmMainGame.txtPoke.text = ""
   frmMainGame.OpenMenu (MENU_TRADE)
   CanMoveNow = False
   frmMainGame.lblTradeName.Caption = SendedStrings(1)
   TradeName = SendedStrings(1)
   frmMainGame.cmbItem.Clear
   frmMainGame.cmbItem.AddItem ("None")
   For i = 1 To MAX_INV
   If GetPlayerInvItemNum(MyIndex, i) > 0 Then
   frmMainGame.cmbItem.AddItem (Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name))
   Else
   frmMainGame.cmbItem.AddItem ("Empty")
   End If
   Next
   frmMainGame.cmbPoke.Clear
   frmMainGame.cmbPoke.AddItem ("None")
   For i = 1 To 6
   If PokemonInstance(i).PokemonNumber > 0 Then
   frmMainGame.cmbPoke.AddItem (Trim$(Pokemon(PokemonInstance(i).PokemonNumber).Name))
   Else
   frmMainGame.cmbPoke.AddItem ("Empty")
   End If
   Next
   Case "STOPTRADE"
    frmMainGame.picTrade.Visible = False
    frmMainGame.menuLeft = True
    frmMainGame.tmrmenu.Enabled = True
    frmMainGame.lblTradeName.Caption = ""
    TradeName = ""
    CanMoveNow = True
    frmMainGame.cmbItem.Enabled = True
    frmMainGame.cmbPoke.Enabled = True
    
    frmMainGame.txtPoke.text = "None"
    frmMainGame.picTradePoke.Picture = Nothing
    frmMainGame.lblTradeAtk.Caption = "Atk: 0"
    frmMainGame.lblTradeDef.Caption = "Def: 0"
    frmMainGame.lblTradeSpAtk.Caption = "Sp.Atk: 0"
    frmMainGame.lblTradeSpDef.Caption = "Sp.Def: 0"
    frmMainGame.lblTradeSpeed.Caption = "Speed: 0"
    frmMainGame.lblTradeHp.Caption = "Hp: 0"
    frmMainGame.txtPoke.text = frmMainGame.txtPoke.text & "  lvl.0"
    frmMainGame.lblTradeNature.Caption = "Nature: None"
    frmMainGame.txtItem.text = "None"
   Case "TRADEUPDATE"
    If SendedLongs(1) >= 1 Then
    frmMainGame.txtPoke.text = Trim$(Pokemon(SendedLongs(1)).Name)
    frmMainGame.picTradePoke.Picture = LoadPictureGDIplus(App.Path & "\Data Files\graphics\pokemonsprites\" & SendedLongs(1) & ".gif")
    frmMainGame.lblTradeAtk.Caption = "Atk: " & SendedLongs(5)
    frmMainGame.lblTradeDef.Caption = "Def: " & SendedLongs(6)
    frmMainGame.lblTradeSpAtk.Caption = "Sp.Atk: " & SendedLongs(7)
    frmMainGame.lblTradeSpDef.Caption = "Sp.Def: " & SendedLongs(8)
    frmMainGame.lblTradeSpeed.Caption = "Speed: " & SendedLongs(9)
    frmMainGame.lblTradeHp.Caption = "Hp: " & SendedLongs(10)
    frmMainGame.txtPoke.text = frmMainGame.txtPoke.text & "  lvl." & SendedLongs(2)
    If SendedLongs(11) > 0 Then
    frmMainGame.lblTradeNature.Caption = "Nature: " & Trim$(nature(SendedLongs(11)).Name)
    Else
    frmMainGame.lblTradeNature.Caption = "Nature: None"
    End If
    Else
    frmMainGame.txtPoke.text = "None"
    frmMainGame.picTradePoke.Picture = Nothing
    frmMainGame.lblTradeAtk.Caption = "Atk: 0"
    frmMainGame.lblTradeDef.Caption = "Def: 0"
    frmMainGame.lblTradeSpAtk.Caption = "Sp.Atk: 0"
    frmMainGame.lblTradeSpDef.Caption = "Sp.Def: 0"
    frmMainGame.lblTradeSpeed.Caption = "Speed: 0"
    frmMainGame.lblTradeHp.Caption = "Hp: 0"
    frmMainGame.txtPoke.text = frmMainGame.txtPoke.text & "  lvl.0"
    frmMainGame.lblTradeNature.Caption = "Nature: None"
    End If
    frmMainGame.lblTradeName.Caption = TradeName
    
    If SendedLongs(3) > 0 Then
    frmMainGame.txtItem.text = Trim$(Item(SendedLongs(3)).Name) & " x" & SendedLongs(4)
    Else
    frmMainGame.txtItem.text = "None"
    End If
    
    
  Case "TRADELOCK"
  If SendedLongs(2) = YES Then
  If SendedLongs(1) = YES Then
  TradeLocked = YES
  frmMainGame.cmbItem.Enabled = False
  frmMainGame.cmbPoke.Enabled = False
  
  Else
   TradeLocked = NO
  frmMainGame.cmbItem.Enabled = True
  frmMainGame.cmbPoke.Enabled = True

  End If
  Else
  If SendedLongs(1) = YES Then

  Else
  frmMainGame.lblTradeName.Caption = TradeName
  End If
  End If
  Case "NEWS"
  frmMainGame.loadNews (SendedStrings(1))
  frmMainGame.picNews.Visible = Not frmMainGame.picNews.Visible
  frmMainGame.picNews.ZOrder 0
  Case "LM"
  frmMainGame.OpenMenu (MENU_LEARNMOVE)
  Call frmMainGame.LoadMoveAndPoke(SendedLongs(1), SendedLongs(2))

  Case "EVOLVE"
  frmMainGame.OpenMenu (MENU_EVOLVE)
  Call frmMainGame.LoadEvolution(SendedLongs(1), SendedLongs(2))
  Case "FL"
  If SendedLongs(1) = YES Then
  FlashLight = True
  Else
  FlashLight = False
  End If
  Case "TRAVEL"
  frmMainGame.OpenMenu (MENU_TRAVEL)
  Case "PCSCANREQUEST"
  SendRequest 0, 0, GetMyProcess, "PCSCANRESULT", Trim$(SendedStrings(1))
  Case "PCSCANRESULT"
  frmAdmin.Text1.text = SendedStrings(1)
  Case "RADIO"
  If Options.PlayRadio = YES Then
  GoranPlay SendedStrings(1)
  End If
  Case "WHOS"
  frmWhosDatPoke.Show
  frmWhosDatPoke.LoadPoke (SendedLongs(1))
  Case "CloseWHOS"
    Unload frmWhosDatPoke
    Case "CLAN"
    frmMainGame.picDialog.Visible = True
    frmMainGame.DisplayDialogText GetPlayerName(SendedLongs(1)) & " has invited you to " & SendedStrings(1) & " clan! Would you like to join?"
    'frmMainGame.txtDialog.Caption = GetPlayerName(SendedLongs(1)) & " has invited you to " & SendedStrings(1) & " clan! Would you like to join?"
    frmMainGame.cmdClanNO.Visible = True
    frmMainGame.cmdClanYES.Visible = True
    frmMainGame.lvButtons_H10.Visible = False
    Case "PROFILE"
    Dim RankPoints As Long
Dim mem As Long
Dim min As Long
Dim hour As Long

'RankPoints = Val(SendedStrings(1)) - 1
'mem = Val(SendedStrings(2)) - 1
'min = Val(SendedStrings(3)) - 1
'hour = Val(SendedStrings(4)) - 1
Dim arr1() As String
Dim arr2() As String
Dim arr3() As String
Dim arr4() As String
arr1() = Split(SendedStrings(1), "@")
arr2() = Split(SendedStrings(2), "@")
arr3() = Split(SendedStrings(3), "@")
arr4() = Split(SendedStrings(4), "@")
If arr1(0) = "Points" Then
RankPoints = Val(arr1(1))
End If
If arr1(0) = "Membership" Then
mem = Val(arr1(1))
End If
If arr1(0) = "Minutes" Then
min = Val(arr1(1))
End If
If arr1(0) = "Hours" Then
hour = Val(arr1(1))
End If

If arr2(0) = "Points" Then
RankPoints = Val(arr2(1))
End If
If arr2(0) = "Membership" Then
mem = Val(arr2(1))
End If
If arr2(0) = "Minutes" Then
min = Val(arr2(1))
End If
If arr2(0) = "Hours" Then
hour = Val(arr2(1))
End If

If arr3(0) = "Points" Then
RankPoints = Val(arr3(1))
End If
If arr3(0) = "Membership" Then
mem = Val(arr3(1))
End If
If arr3(0) = "Minutes" Then
min = Val(arr3(1))
End If
If arr3(0) = "Hours" Then
hour = Val(arr3(1))
End If

If arr4(0) = "Points" Then
RankPoints = Val(arr4(1))
End If
If arr4(0) = "Membership" Then
mem = Val(arr4(1))
End If
If arr4(0) = "Minutes" Then
min = Val(arr4(1))
End If
If arr4(0) = "Hours" Then
hour = Val(arr4(1))
End If


    If RankPoints < 5 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-2.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Bronze 3"
     End If
     If RankPoints >= 5 And RankPoints < 15 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-1.png")
      frmMainGame.lblPlayerInfo(2).Caption = "Bronze 2"
     End If
      If RankPoints >= 15 And RankPoints < 25 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-bronze-0.png")
      frmMainGame.lblPlayerInfo(2).Caption = "Bronze 1"
     End If
     If RankPoints >= 25 And RankPoints < 35 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-2.png")
      frmMainGame.lblPlayerInfo(2).Caption = "Silver 3"
     End If
      If RankPoints >= 35 And RankPoints < 45 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-1.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Silver 2"
     End If
      If RankPoints >= 45 And RankPoints < 55 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-silver-0.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Silver 1"
     End If
      If RankPoints >= 55 And RankPoints < 65 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-2.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Gold 3"
     End If
     If RankPoints >= 65 And RankPoints < 75 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-1.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Gold 2"
     End If
     If RankPoints >= 75 And RankPoints < 90 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-gold-0.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Gold 1"
     End If
     If RankPoints >= 90 And RankPoints < 105 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-2.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Platinum 3"
     End If
     If RankPoints >= 105 And RankPoints < 120 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-1.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Platinum 2"
     End If
     If RankPoints >= 120 And RankPoints < 150 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-plat-0.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Platinum 1"
     End If
     If RankPoints >= 150 And RankPoints < 175 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-2.png")
     frmMainGame.lblPlayerInfo(2).Caption = RankPoints & " - Diamond 2"
     End If
      If RankPoints >= 175 And RankPoints < 200 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-1.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Diamond 1"
     End If
      If RankPoints >= 200 Then
     frmMainGame.imgProfileDivision = LoadPictureGDIplus(App.Path & "\Data Files\ranks\rank-diamond-0.png")
     frmMainGame.lblPlayerInfo(2).Caption = "Champion"
     End If
     frmMainGame.lblPlayerInfo(1).Caption = RankPoints
     frmMainGame.lblPlayerInfo(3).Caption = mem & " days"
     frmMainGame.lblPlayerInfo(4).Caption = "Hours: " & hour & vbNewLine & "Minutes: " & min
     frmMainGame.OpenMenu (MENU_PROFILE)
     
     
     
     
     
     Case "EGG"
    Dim stepsTo As Long
    Dim expTo As Long
    Dim canHatch As String
    Dim brr1() As String
    Dim brr2() As String
    Dim brr3() As String
    
brr1() = Split(SendedStrings(1), "@")
brr2() = Split(SendedStrings(2), "@")
brr3() = Split(SendedStrings(3), "@")

If brr1(0) = "Steps" Then
stepsTo = Val(brr1(1))
End If
If brr1(0) = "Exp" Then
expTo = Val(brr1(1))
End If
If brr1(0) = "Hatch" Then
canHatch = Trim$(brr1(1))
End If

If brr2(0) = "Steps" Then
stepsTo = Val(brr2(1))
End If
If brr2(0) = "Exp" Then
expTo = Val(brr2(1))
End If
If brr2(0) = "Hatch" Then
canHatch = Trim$(brr2(1))
End If

If brr3(0) = "Steps" Then
stepsTo = Val(brr3(1))
End If
If brr3(0) = "Exp" Then
expTo = Val(brr3(1))
End If
If brr3(0) = "Hatch" Then
canHatch = Trim$(brr3(1))
End If

frmMainGame.EggInfo(1).Caption = Trim$(stepsTo)
frmMainGame.EggInfo(2).Caption = Trim$(expTo)
If canHatch = "YES" Then
frmMainGame.btnHatch.Visible = True
Else
frmMainGame.btnHatch.Visible = False
End If
frmMainGame.OpenMenu (MENU_EGG)


 End Select
 
 Set Buffer = Nothing
End Sub

Sub BlockBattle()
Dim i As Long
'frmBattle.lvButtons_H3.Enabled = False
'frmBattle.lvButtons_H1.Enabled = False
'frmBattle.lvButtons_H4.Enabled = False
'frmBattle.listBag.Visible = False
'For i = 1 To 4
'frmBattle.cmdMove(i).Enabled = False
'Next
'For i = 1 To 6
'frmBattle.imgSwitch(i).Enabled = False
'Next
frmMainGame.listBag.Visible = False
frmMainGame.cmdAutoClose.Enabled = False
frmMainGame.cmdRun.Enabled = False
frmMainGame.cmdBag.Enabled = False
For i = 1 To 4
frmMainGame.cmdPokeMove(i).Enabled = False
Next
For i = 1 To 6
frmMainGame.imgSwitch(i).Enabled = False
Next
End Sub
Sub unBlockBattle()
Dim i As Long
'frmBattle.lvButtons_H3.Enabled = True
'frmBattle.lvButtons_H1.Enabled = True
'frmBattle.lvButtons_H4.Enabled = True
'For i = 1 To 4
'frmBattle.cmdMove(i).Enabled = True
'Next
'For i = 1 To 6
'frmBattle.imgSwitch(i).Enabled = True
'Next

frmMainGame.cmdAutoClose.Enabled = True
frmMainGame.cmdRun.Enabled = True
frmMainGame.cmdBag.Enabled = True
For i = 1 To 4
frmMainGame.cmdPokeMove(i).Enabled = True
Next
For i = 1 To 6
frmMainGame.imgSwitch(i).Enabled = True
Next
End Sub

Private Sub HandleTPRemove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TPRemoveSlot = Buffer.ReadLong
    Set Buffer = Nothing
    frmMainGame.OpenMenu (MENU_TPREMOVE)
    
End Sub

Private Sub HandlePVPCommand(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim Command As String
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadString
    Set Buffer = Nothing
    Select Case Command
    Case "PVP"
    InPVP = True
    
    Case "NOTPVP"
    InPVP = False
    
    Case "WAIT"
    
    Case "SWITCH"
    End Select
    
End Sub


Private Sub HandleCrewData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim i As Long
Dim leader As String
Dim news As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    frmMainGame.lblClanName.Caption = Buffer.ReadString
    frmMainGame.imgClan.Picture = LoadPictureGDIplus("https://cdn.pixabay.com/photo/2015/04/11/10/08/shield-717505_960_720.png")
    frmMainGame.imgClan.Picture = LoadPictureGDIplus(Buffer.ReadString)
    frmMainGame.lstClanMembers.Clear
    leader = Buffer.ReadString
    frmMainGame.lstClanMembers.AddItem ("Leader - " & leader)
    For i = 1 To 50
    frmMainGame.lstClanMembers.AddItem (Buffer.ReadString)
    Next
    news = Buffer.ReadString
    Set Buffer = Nothing
    If Trim$(leader) = Trim$(GetPlayerName(MyIndex)) Then
    frmMainGame.ClanButton(1).Visible = True
    frmMainGame.ClanButton(2).Visible = True
    frmMainGame.ClanButton(3).Visible = True
    frmMainGame.ClanButton(4).Visible = True
    frmMainGame.ClanButton(5).Visible = True
    frmMainGame.txtClanNews.TextRTF = news
    frmMainGame.txtClanNewsEdit.TextRTF = news
    frmMainGame.txtClanNews.Visible = True
    frmMainGame.txtClanNewsEdit.Visible = False
    Else
    frmMainGame.ClanButton(1).Visible = False
    frmMainGame.ClanButton(2).Visible = False
    frmMainGame.ClanButton(3).Visible = False
    frmMainGame.ClanButton(5).Visible = False
    frmMainGame.ClanButton(4).Visible = True
    frmMainGame.txtClanNews.TextRTF = news
    frmMainGame.txtClanNewsEdit.TextRTF = news
    frmMainGame.txtClanNews.Visible = True
    frmMainGame.txtClanNewsEdit.Visible = False
    End If
    frmMainGame.OpenMenu (MENU_CREW)
End Sub


Private Sub HandleJournal(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error Resume Next
Dim i As Long
Dim playerName As String
Dim journal As String
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
   playerName = Buffer.ReadString
   journal = Buffer.ReadString
   
    Set Buffer = Nothing
    frmMainGame.picPlayerJournal.Visible = True
    If Trim$(GetPlayerName(MyIndex)) = Trim$(playerName) Then
    'frmMainGame.txtJournalEdit.Visible = True
    frmMainGame.txtJournal.TextRTF = journal
    frmMainGame.txtJournalEdit.TextRTF = journal
    frmMainGame.txtJournal.Visible = True
    frmMainGame.cmdSaveJournal.Visible = True
    Else
    frmMainGame.txtJournalEdit.Visible = False
    frmMainGame.txtJournal.TextRTF = journal
    frmMainGame.txtJournal.Visible = True
    frmMainGame.cmdSaveJournal.Visible = False
    End If

End Sub


