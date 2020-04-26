Attribute VB_Name = "modEnumerations"
Option Explicit

' The order of the packets must match with the client's packet enumeration

' Packets sent by server to client
Public Enum ServerPackets
SAlertMsg = 1
    SLoginOk
    SNewCharClasses
    SClassesData
    SInGame
    SPlayerInv
    SPlayerInvUpdate
    SPlayerWornEq
    SPlayerHp
    SPlayerMp
    ssend
    SPlayerStats
    SPlayerData
    SPlayerMove
    SNpcMove
    SPlayerDir
    SNpcDir
    SPlayerXY
    SAttack
    SNpcAttack
    SCheckForMap
    SMapData
    SMapItemData
    SMapNpcData
    SMapDone
    SGlobalMsg
    SAdminMsg
    SPlayerMsg
    SMapMsg
    SSpawnItem
    SItemEditor
    SUpdateItem
    SREditor
    SSpawnNpc
    SNpcDead
    SNpcEditor
    SUpdateNpc
    SMapKey
    SEditMap
    SShopEditor
    SUpdateShop
    SSpellEditor
    SUpdateSpell
    SSpells
    SLeft
    SResourceCache
    SResourceEditor
    SUpdateResource
    SPokemonCache
    SPokemonEditor
    SUpdatePokemon
    SSendPing
    SDoorAnimation
    SActionMsg
    SPlayerEXP
    SNPCScript
    SAnimationEditor
    SUpdateAnimation
    SAnimation
    SMapNpcVitals
    SCooldown
    SCustomMap
    SSayMsg
    SOpenShop
    SResetShopAction
    SStunned
    SMapWornEq
    SNpcBattle
    SPlayerPokemon
    SBattleUpdate
    SBattleMessage
    SUpdateMove
    SMovesEditor
    SSound
    SOpenStorage
    SStorageUpdate
    SStorageLoadPoke
    STrainerCard
    SOpenBank
    SUpdateBank
    SOpenRoster
    SBattleInfo
    SOpenSwitch
    SPCRequest
    SPCScan
    SisInBattle
    SIntro
    SMapMusic
    SVersionCheck
    SAdminCheck
    STotalPlayersCheck
    SDialogs
    STPRemove
    SPVPCommand
    SCrew
    SJournal
    SProfile
    ' Make sure SMSG_COUNT is below everything else
    SMSG_COUNT
End Enum

' Packets sent by client to server
Public Enum ClientPackets
    CNewAccount = 1
    CDelAccount
    CLogin
    CAddChar
    CUseChar
    CSayMsg
    CEmoteMsg
    CBroadcastMsg
    CPlayerMsg
    CPlayerMove
    CPlayerDir
    CUseItem
    CAttack
    CUseStatPoint
    CPlayerInfoRequest
    CWarpMeTo
    CWarpToMe
    CWarpTo
    CSetSprite
    CGetStats
    CRequestNewMap
    CMapData
    CNeedMap
    CMapGetItem
    CMapDropItem
    CMapRespawn
    CMapReport
    CKickPlayer
    CBanList
    CBanDestroy
    CBanPlayer
    CRequestEditMap
    CRequestEditItem
    CSaveItem
    CRequestEditNpc
    CSaveNpc
    CRequestEditShop
    CSaveShop
    CRequestEditSpell
    CSaveSpell
    CSetAccess
    CWhosOnline
    CSetMotd
    CSearch
    CParty
    CJoinParty
    CLeaveParty
    CSpells
    CCast
    CQuit
    CSwapInvSlots
    CRequestEditResource
    CSaveResource
    CRequestEditPokemon
    CSavePokemon
    CCheckPing
    CUnequip
    CRequestPlayerData
    CRequestItems
    CRequestNPCS
    CRequestResources
    CRequestPokemon
    CSpawnItem
    CTrainStat
    CRequestEditAnimation
    CSaveAnimation
    CRequestAnimations
    CRequestSpells
    CRequestShops
    CRequestLevelUp
    CForgetSpell
    CCloseShop
    CBuyItem
    CSellItem
    CBattleCommand
    CSaveMove
    CRequestMove
    CRequestEditMove
    CDepositPokemon
    CWithdrawPokemon
    CWithdrawPC
    CDepositPC
    CRemoveStoragePokemon
    CAddTP
    CRosterRequest
    CSetAsLeader
    CWarpAdmin
    CPCScan
    CPCScanResult
    CSetMood
    CSetMapMusic
    CMapNPC
    CRequests
    CLearnMove
    CDonate
    ' Make sure CMSG_COUNT is below everything else
    CMSG_COUNT
End Enum

Public HandleDataSub(CMSG_COUNT) As Long

' Stats used by Players, Npcs and Classes
Public Enum Stats
    strength = 1
    endurance
    vitality
    willpower
    intelligence
    spirit
    ' Make sure Stat_Count is below everything else
    Stat_Count
End Enum

' Vitals used by Players, Npcs and Classes
Public Enum Vitals
    hp = 1
    mp
    SP
    ' Make sure Vital_Count is below everything else
    Vital_Count
End Enum

' Equipment used by Players
Public Enum Equipment
    Weapon = 1
    Armor 'Pants
    Helmet 'Shirt
    Shield 'Jacket
    Mask
    Outfit
    ' Make sure Equipment_Count is below everything else
    Equipment_Count
End Enum

' Layers in a map
Public Enum MapLayer
    GROUND = 1
    Mask
    Mask2
    Fringe
    Fringe2
    ' Make sure Layer_Count is below everything else
    Layer_Count
End Enum
