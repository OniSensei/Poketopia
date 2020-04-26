Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public map As MapRec
Public UpMap As MapRec
Public DownMap As MapRec
Public RightMap As MapRec
Public LeftMap As MapRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Pokemon(1 To MAX_POKEMONS) As PokemonRec
Public PokemonMove(1 To MAX_MOVES) As MoveRec
Public nature(1 To MAX_NATURES) As NatureRec
' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public EditPokemons(1 To MAX_MAP_POKEMONS) As MapPokemonRec
Public Weather(1 To MAX_MAPS) As WeatherRec

' options
Public Options As OptionsRec


Private Type MoveRec
Name As String * 20
Type As String * 20
demage As Long
pp As Long
power As Long
accuracy As Long
Description As String * 200
effect As Long
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

' Type recs
Private Type OptionsRec
    SavePass As Byte
    Password As String * NAME_LENGTH
    Username As String * NAME_LENGTH
    IP As String
    Port As Long
    music As String
    repeatmusic As Long
    PlayMusic As Long
    CameraFollowPlayer As Long
    FormTransparency As Long
    PlayRadio As Long
    NearbyMaps As Long
End Type

Private Type PokemonMoveRec
movenumber As Long
CPP As Long
End Type

Public Type BattleMoveRec
number As Long
pp As Long
power As Long
accuracy As Long
End Type
 
Public Type PokemonBattleEnemyRec
PokemonNumber As Long
MaxHp As Long
HP As Long
ATK As Long
DEF As Long
SPD As Long
SPATK As Long
SPDEF As Long
Level As Long
isFreezed As Long
isBurned As Long
isParalized As Long
isPoisoned As Long
isSleeping As Long
isAttracted As Long
isConfused As Long
isCursed As Long
moves(1 To 4) As BattleMoveRec
status As Long
nature As Long
isShiny As Long
End Type
 
 


Public Type PokemonInstanceRec
    PokemonNumber As Long
    ' stats
    HP As Long
    pp As Long
    moves(1 To 4) As BattleMoveRec
    ATK As Long
    DEF As Long
    SPD As Long
    SPATK As Long
    SPDEF As Long
    ' stuff
    Happiness As Long
    EXP As Long
    Level As Long
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
status As Long
expNeeded As Long
isShiny As Long
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
    MaxPP As Long
    moves(1 To 30) As Integer
    movesLV(1 To 30) As Integer
    Type As Byte
    Type2 As Byte
    Evolution As Long
    EvolvesTo As Long
    ATK As Long
    DEF As Long
    SPD As Long
    SPATK As Long
    SPDEF As Long
    Rareness As Long
    BaseEXP As Long
    PercentFemale As Long
    Happiness As Long
    CatchRate As Long
    Stone As String * 40
End Type

Public Type PlayerInvRec
    num As Byte
    Value As Long
End Type

Private Type SpellAnim
    spellnum As Integer
    Timer As Long
    FramePointer As Long
End Type

Private Type PlayerRec
    ' General
    Name As String
    Class As Byte
    Sprite As Integer
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.stat_count - 1) As Byte
    POINTS As Byte
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Byte
    ' Position
    map As Integer
    x As Byte
    y As Byte
    dir As Byte
    ' Client use only
    MaxHp As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    SMap As Long
    SX As Long
    SY As Long
    Bedages(1 To MAX_GYMS) As Long
    StoredPC As Long
    Badge(1 To MAX_GYMS) As Long
    inBattle As Long
    mood As Long
    Pokes(1 To 6) As Long
    isMember As Long
    notVisible As Boolean
    HasBike As Long
End Type

Private Type TileDataRec
    x As Byte
    y As Byte
    tileset As Byte
End Type

Public Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Type As Byte
    data1 As Long
    data2 As Long
    data3 As Long
End Type


Private Type NPCScriptRec
   Name As String * 30
   script As String
End Type


Public Type MapPokemonRec
    PokemonNumber As Long
    Chance As Long
    LevelFrom As Long
    LevelTo As Long
    Custom As Long
    ATK As Long
    DEF As Long
    SPATK As Long
    SPDEF As Long
    SPD As Long
    HP As Long
    nature As Long
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Revision As Long
    Moral As Byte
    tileset As Long
    
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    
    music As Byte
    
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
    Stat(1 To Stats.stat_count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    pic As Integer

    Type As Byte
    data1 As Integer
    data2 As Integer
    data3 As Integer
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.stat_count - 1) As Byte
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.stat_count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    CatchRate As Long
    AddHP As Long
End Type

Private Type MapItemRec
    num As Byte
    Value As Long
    frame As Byte
    x As Byte
    y As Byte
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100
    Sprite As Integer
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    DropChance As Integer
    DropItem As Byte
    DropItemValue As Integer
    Stat(1 To Stats.stat_count - 1) As Byte
    faction As Byte
    HP As Long
    EXP As Long
    Animation As Long
    CanMove As Long
    Paperdoll1 As Long
    Paperdoll2 As Long
    Paperdoll3 As Long
End Type

Private Type MapNpcRec
    num As Byte
    Target As Byte
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    map As Integer
    x As Byte
    y As Byte
    dir As Byte
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    script As String
End Type

Private Type TradeItemRec
    Item As Long
    itemvalue As Long
    CostItem As Long
    CostValue As Long
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
    dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Long
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    x As Long
    y As Long
    ResourceState As Byte
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
    Health As Byte
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
End Type

Private Type ActionMsgRec
    message As String
    Created As Long
    Type As Long
    color As Long
    Scroll As Long
    x As Long
    y As Long
    Width As Long
    Height As Long
    backColor As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    x As Long
    y As Long
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    x As Long
    y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    FrameIndex(0 To 1) As Long
End Type

Private Type WeatherRec
PicName As String * 20
Pics As Long
pic_rotation(1 To 100) As Long
pics_x(1 To 100) As Long
pics_Y(1 To 100) As Long
speed As Long
End Type

'Evilbunnie's DrawnChat system
Public Chat(1 To 20) As ChatRec

'Evilbunnie's DrawnChat system
Private Type ChatRec
    text As String
    Colour As Long
End Type
