Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public LevelExp(1 To 100) As LevelExpRec
Public map(1 To MAX_MAPS) As MapRec
Public MapCache(1 To MAX_MAPS) As Cache
Public TempTile(1 To MAX_MAPS) As TempTileRec
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public player(1 To MAX_PLAYERS) As PlayerRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public Class() As ClassRec
Public item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Pokemon(1 To MAX_POKEMONS) As PokemonRec
Public PokemonMove(1 To MAX_MOVES) As MoveRec
Public nature(1 To MAX_NATURES) As NatureRec
Public Options As OptionsRec
Public Types(1 To 18) As TypesRec
Public MoveTypes(1 To MAX_MOVES) As Byte
'Setting leader



'TYPES

Public Type StatsRec
hp As Long
atk As Long
def As Long
spd As Long
spatk As Long
spdef As Long
criticalHit As Long
accuracy As Long
End Type

Private Type MoveEffectRec
lastMove As String
state As String
move As String
statusReturn As Long
moveReturn As String
hpDef As Long
atkdef As Long
defDef As Long
spAtkDef As Long
spDefDef As Long
speedDef As Long
hpAtk As Long
atkAtk As Long
defAtk As Long
spAtkAtk As Long
spDefAtk As Long
speedAtk As Long
paydayused As Long
attack As Long
messgae As String
PokemonNumber As Long
resetStatsMe As Long 'BOOL
resetStatsMeTo As Long 'RESET TO
resetStats As Long 'BOOL
resetStatsTo As Long 'RESET TO
onlyOneMove As Long 'Bool
onlyOneMoveRounds As Long
onlyOneMoveAdditionalEffect As String
End Type

Private Type moveUsageTempRec
defenderStatus As Long
defenderStatusRounds As Long
attackerStatus As Long
attackerStatusRounds As Long
powerModifier As Long
resetStatsMe As Long
resetStatsMeTo As Long
resetStats As Long
resetStatsTo As Long
onlyOneMove As Boolean
onlyOneMoveRounds As Long
onlyOneMoveAdditionalEffect As String
damageInflict As Long
effectBegin As String
effecrLast As Long
HPModifier As Long
HPDamageModifier As Long
HPTotalModifier As Long
powerSet As Long
isCritical As Long
customAttackUsage As String
recoilHPTotalModifier As Long
recoilHPDamageModifier As Long
recoilHPCurrentModifier As Long
fleeBattle As Boolean
multiHit As Long
End Type


Private Type TypesRec
NORMAL As Double
FIGHT As Double
FLYING As Double
POISON As Double
GROUND As Double
ROCK As Double
BUG As Double
GHOST As Double
STEEL As Double
FIRE As Double
WATER As Double
GRASS As Double
ELECTRIC As Double
PSYCHIC As Double
ICE As Double
DRAGON As Double
DARK As Double
FAIRY As Double
End Type

Private Type LevelExpRec
Erratic As Long
Fast As Long
Medium_Fast As Long
Medium_Slow As Long
Slow As Long
Fluctuating As Long
End Type

Private Type MoveRec
Name As String * 20
Type As String * 20
demage As Long
pp As Long
power As Long
accuracy As Long
Description As String * 200
Effect As Long
Category As String * 20
Generation As Integer
doesDamageIfMiss As Integer
missDamageModifier As Integer
isFlinching As Integer
flinchChances As Integer
isCharging As Integer
canBeHitOnCharging As Integer
chargeFirst As Integer
chargeTurns As Integer
critical_hit_ration As Integer
isMultiTurn As Integer
multiTurnLowerLimit As Integer
multiTurnUpperLimit As Integer
isMultiHit As Integer
multiHitLowerLimit As Integer
multiHitUpperLimit As Integer
isRecoil As Integer
isHealing As Integer
priority As Integer
recoilModifier As Integer
isAttackerStatChanging As Integer
attackerStatChangeModifier As Integer
attackerStatChangeIndex As String * 20
attackerStatChangeChances As Integer
isOpponentStatChanging As Integer
opponentStatChangeModifier As Integer
opponentStatChangeIndex As String * 20
opponentStatChangeChances As Integer
isOpponentNonVolatileStatusInducing As Integer
opponentNonVolatileStatusInducingChances As Integer
nonVolatileStatusType As String * 20
isOpponentStatResetting As Integer
opponentStatResettingChances As Integer
isAttackerStatResetting As Integer
attackerStatResettingChances As Integer
isHpRestoring As Integer
HpRestoringChances As Integer
hpRestoreModifier As Integer
isOpponentVolatileStatusIndiucing As Integer
opponentVolatileStatusInducingChances As Integer
OpponentVolatileStatusType As String * 20
isAttackerVolatileStatusInducing As Integer
AttackerVolatileStatusType As String * 20
attackerVolatileStatusInducingChances As Integer
InteralType1 As String * 20
InteralType2 As String * 20
End Type
 

 
Private Type BattleMoveRec
number As Long
pp As Long
power As Long
accuracy As Long
End Type
 
Private Type PokemonBattleEnemyRec
PokemonNumber As Long
MaxHp As Long
hp As Long
atk As Long
def As Long
spd As Long
spatk As Long
spdef As Long
level As Long
isFreezed As Long
isBurned As Long
isParalized As Long
isPoisoned As Long
isSleeping As Long
isAttracted As Long
isConfused As Long
isCursed As Long
isFlinched As Long
moves(1 To 4) As BattleMoveRec
MapSlot As Long
status As Long
nature As Long
statusturn As Long
turnsneed As Long
isShiny As Long
statsChange As StatsRec
FirstMove As Long
End Type
 
 
Public Type PokemonInstanceRec
    PokemonNumber As Long
    ' stats
    hp As Long
    pp As Long
    moves(1 To 4) As BattleMoveRec
    atk As Long
    def As Long
    spd As Long
    spatk As Long
    spdef As Long
    ' stuff
    Happiness As Long
    EXP As Long
    level As Long
    Sex As Byte
    nature As Long
    MaxHp As Long
    TP As Long
    isFreezed As Long
isBurned As Long
isParalized As Long
isPoisoned As Long
isSleeping As Long
isAttracted As Long
isConfused As Long
isCursed As Long
isFlinched As Long
status As Long
batk As Long
bdef As Long
bspd As Long
bspatk As Long
bspdef As Long
statusstun As Long
turnsneed As Long
isShiny As Long
isTradeable As Long
HoldingItem As Long
End Type

Public Type NatureRec
     Name As String * 20
     AddHP As Long
     AddAtk As Long
     AddDef As Long
     AddSpd As Long
     AddSpAtk As Long
     AddSpDef As Long
End Type

Private Type PokemonRec
    Name As String * NAME_LENGTH
    MaxHp As Long
    maxpp As Long
    moves(1 To 30) As Integer
    movesLV(1 To 30) As Integer
    Type As Byte
    Type2 As Byte
    Evolution As Long
    EvolvesTo As Long
    atk As Long
    def As Long
    spd As Long
    spatk As Long
    spdef As Long
    Rareness As Long
    BaseEXP As Long
    PercentFemale As Long
    Happiness As Long
    CatchRate As Long
    Stone As String * 40
End Type

Private Type OptionsRec
    Game_Name As String
    MOTD As String
    Port As Long
    Website As String
End Type

Private Type PlayerInvRec
    Num As Byte
    value As Long
End Type

Private Type Cache
    Data() As Byte
End Type

Public Type PlayerRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    
    ' General
    Name As String * NAME_LENGTH
    Sex As Byte
    Class As Byte
    Sprite As Integer
    level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Byte
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Byte
    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    ' Position
    map As Integer
    x As Byte
    y As Byte
    Dir As Byte
    lastx As Byte
    lastty As Byte
    lastDir As Byte
    ' pokemons
    PokemonInstance(1 To 6) As PokemonInstanceRec
    StoragePokemonInstance(1 To 250) As PokemonInstanceRec
    SMap As Long
    SX As Long
    SY As Long
    Bedages(1 To MAX_GYMS) As Long
    StoredPC As Long
    mood As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    buffer As clsBuffer
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    PartyPlayer As Long
    InParty As Byte
    TargetType As Byte
    Target As Byte
    PartyStarter As Byte
    GettingMap As Byte
    SpellBuffer As Long
    SpellBufferTimer As Long
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    TradeName As String * NAME_LENGTH
    isTrading As Long
    TradePoke As Long
    TradeItem As Long
    TradeItemVal As Long
    TradeLocked As Long
    ' pokemon!
    PokemonBattle As PokemonBattleEnemyRec
    BattleType As Byte
    BattleTurn As Boolean
    BattleCurrentTurn As Long
    WaitingForSwitch As Long
    isInBattle As Long
    isCatchable As Long
    lastPoke As Long
    moveData As MoveEffectRec
    statChanges(1 To 6) As StatsRec
    moveUsageTemp As moveUsageTempRec
    inNPCBattle As Boolean
    NPCBattlePokemons(1 To 6) As PokemonBattleEnemyRec
    NPCBattle As Long
    NPCBattleSelectedPoke As Long
    NPCBattlePokesAvailable As Long
    LearnMoveNumber As Long
    LearnMovePokemon As Long
    LearnMovePokemonName As String
    FishingTimer As Long
    CanFish As Boolean
    CanFishOnSpot As Boolean
    FishSpotX As Long
    FishSpotY As Long
    SpecialEvolveSlot As Long
    SpecialEvolveTo As Long
    SpecialEvolveItem As Long
    LearnMoveIsTM As Boolean
    isInTPRemoval As Boolean
    'PVP
    isInPVP As Boolean
    PVPEnemy As String
    PVPCommandUsed As Long
    PVPCommandNum As Long
    PVPHasUsed As Boolean
    PVPSlot As Long
    PVPTurnAdvantage As Boolean
    clanInvite As Boolean
    clanInviteIndex As Long
    notVisible As Boolean
    'Dialog
    hasDialogTrigger As Boolean
    dialogTriggerData1 As Long
    dialogTriggerData2 As Long
    dialogTriggerData3 As String
    'Egg
    eggStepsTemp As Long
    eggExpTemp As Long
    'Bike
    HasBike As Long
End Type

Private Type TileDataRec
    x As Byte
    y As Byte
    Tileset As Byte
End Type

Private Type NPCScriptRec
   Name As String * 30
   script As String
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
End Type




Private Type MapPokemonRec
    PokemonNumber As Long
    Chance As Long
    LevelFrom As Long
    LevelTo As Long
    Custom As Long
    atk As Long
    def As Long
    spatk As Long
    spdef As Long
    spd As Long
    hp As Long
    nature As Long
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    Tileset As Long
    
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    
    Music As Byte
    
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    
    MaxX As Byte
    MaxY As Byte
    Pokemon(1 To MAX_MAP_POKEMONS) As MapPokemonRec
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Integer
    MusicName As String * 20
    NPCScripts(1 To MAX_MAP_NPCS) As String
End Type



Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Pic As Integer

    Type As Byte
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    Speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    CatchRate As Long
    AddHP As Long
End Type

Private Type MapItemRec
    Num As Byte
    value As Long
    x As Byte
    y As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sprite As Integer
    SpawnSecs As Long
    Behaviour As Byte
    range As Byte
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    Stat(1 To Stats.Stat_Count - 1) As Byte
    faction As Byte
    hp As Long
    EXP As Long
    Animation As Long
    CanMove As Long
    Paperdoll1 As Long
    Paperdoll2 As Long
    Paperdoll3 As Long
End Type

Private Type MapNpcRec
    Num As Integer
    Target As Integer
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    x As Byte
    y As Byte
    Dir As Integer
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    script As String
End Type

Private Type TradeItemRec
    item As Long
    ItemValue As Long
    costitem As Long
    costvalue As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Type As Byte
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    map As Long
    x As Long
    y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    range As Long
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
End Type

Private Type TempTileRec

    DoorOpen() As Byte
    DoorTimer As Long
End Type

Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    x As Long
    y As Long
    cur_health As Byte
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    ResourceType As Byte
    ResourceImage As Byte
    ExhaustedImage As Byte
    ItemReward As Long
    ToolRequired As Long
    health As Byte
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type




