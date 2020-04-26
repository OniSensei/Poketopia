Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"
'Version
Public Const VersionCode As Long = 1133
Public Const TCP_CODE As Long = 432001
Public Const BATTLE_NO As Long = 5

' Version constants
Public Const CLIENT_MAJOR As Byte = 4
Public Const CLIENT_MINOR As Byte = 0
Public Const CLIENT_REVISION As Byte = 0
Public Const MAX_LINES As Integer = 500 ' Used for frmServer.txtText
Public Const GoldNeeded As Boolean = False

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants

Public Const MAX_PLAYERS As Long = 500
Public Const MAX_ITEMS As Byte = 255
Public Const MAX_NPCS As Byte = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Byte = 35
Public Const MAX_MAP_ITEMS As Byte = 255
Public Const MAX_MAP_NPCS As Byte = 30
Public Const MAX_SHOPS As Byte = 50
Public Const MAX_PLAYER_SPELLS As Byte = 20
Public Const MAX_SPELLS As Byte = 255
Public Const MAX_TRADES As Byte = 20
Public Const MAX_RESOURCES As Byte = 100
Public Const MAX_LEVELS As Byte = 100
Public Const MAX_POKEMONS As Long = 721
Public Const MAX_GYMS As Byte = 20
Public Const MAX_MOVES As Long = 621
Public Const MAX_NATURES As Long = 71
' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' Account constants
Public Const NAME_LENGTH As Byte = 20

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 1000
Public Const MAX_MAPX As Byte = 17
Public Const MAX_MAPY As Byte = 11
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAX_MAP_POKEMONS As Long = 30
' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BATTLE As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_SPAWN As Byte = 13
Public Const TILE_TYPE_STORAGE As Byte = 14
Public Const TILE_TYPE_BANK As Byte = 15
Public Const TILE_TYPE_GYMBLOCK As Byte = 16
Public Const TILE_TYPE_CUSTOMSCRIPT As Byte = 17

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_POTIONADDHP As Byte = 5
Public Const ITEM_TYPE_POTIONADDMP As Byte = 6
Public Const ITEM_TYPE_POTIONADDSP As Byte = 7
Public Const ITEM_TYPE_POTIONSUBHP As Byte = 8
Public Const ITEM_TYPE_POTIONSUBMP As Byte = 9
Public Const ITEM_TYPE_POTIONSUBSP As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13
Public Const ITEM_TYPE_POKEPOTION As Byte = 14
Public Const ITEM_TYPE_POKEBALL As Byte = 15
Public Const ITEM_TYPE_SCRIPT As Byte = 16
Public Const ITEM_TYPE_STONE As Byte = 17
Public Const ITEM_TYPE_MASK As Byte = 18
Public Const ITEM_TYPE_OUTFIT As Byte = 19
Public Const ITEM_TYPE_HOLDING As Byte = 20
' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2

' ********************************************
' Default starting location [Server Only]
Public Const START_MAP As Integer = 115
Public Const START_X As Integer = 10
Public Const START_Y As Integer = 9

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

'POKEMON TYPES
Public Const TYPE_NONE As Long = 0
Public Const TYPE_NORMAL As Long = 1
Public Const TYPE_FIGHTING As Long = 2
Public Const TYPE_FLYING As Long = 3
Public Const TYPE_POISON As Long = 4
Public Const TYPE_GROUND As Long = 5
Public Const TYPE_ROCK As Long = 6
Public Const TYPE_BUG As Long = 7
Public Const TYPE_GHOST As Long = 8
Public Const TYPE_STEEL As Long = 9
Public Const TYPE_FIRE As Long = 10
Public Const TYPE_WATER As Long = 11
Public Const TYPE_GRASS As Long = 12
Public Const TYPE_ELECTRIC As Long = 13
Public Const TYPE_PSYCHIC As Long = 14
Public Const TYPE_ICE As Long = 15
Public Const TYPE_DRAGON As Long = 16
Public Const TYPE_DARK As Long = 17
Public Const TYPE_FAIRY As Long = 18
'Stats
Public Const STAT_HP As Long = 1
Public Const STAT_ATK As Long = 2
Public Const STAT_DEF As Long = 3
Public Const STAT_SPATK As Long = 4
Public Const STAT_SPDEF As Long = 5
Public Const STAT_SPEED As Long = 6
'Gym
Public Const GYM_DEFEATED As Long = 1
Public Const GYM_UNDEFEATED As Long = 0


'Pokemon status
Public Const STATUS_NOTHING As Long = 1
Public Const STATUS_FREEZED As Long = 2
Public Const STATUS_BURNED As Long = 3
Public Const STATUS_PARALIZED As Long = 4
Public Const STATUS_POISONED As Long = 5
Public Const STATUS_SLEEPING As Long = 6
Public Const STATUS_ATTRACTED As Long = 7
Public Const STATUS_CONFUSED As Long = 8
Public Const STATUS_CURSED As Long = 9
Public Const STATUS_FLINCHED As Long = 10
Public Const STATUS_BADLYPOISONED As Long = 11
'///////////////////GOLF(GORAN)//////////////////
'//////////////////MOVES////////////////////////
'/////////////////EFFECTS//////////////////////
Public Const EFFECT_DEALS_DEMAGE As Byte = 0
Public Const EFFECT_SLEEP As Byte = 1
Public Const EFFECT_POISON As Byte = 2
Public Const EFFECT_GAIN_HALF_HP As Byte = 3
Public Const EFFECT_BURN As Byte = 4
Public Const EFFECT_FREEZE As Byte = 5
Public Const EFFECT_PARALYZE As Byte = 6
Public Const EFFECT_DEMAGE_HELVED As Byte = 7
Public Const EFFECT_ONLY_SLEEP_GAIN_HALF_HP As Byte = 8
Public Const EFFECT_FLYING_UNUSEFULL As Byte = 9
Public Const EFFECT_ATTACK1 As Byte = 10
Public Const EFFECT_DEFENSE1 As Byte = 11
Public Const EFFECT_SP1 As Byte = 12
Public Const EFFECT_EVASIVENESS1 As Byte = 13
Public Const EFFECT_CANNOT_BE_EVADED As Byte = 14
Public Const EFFECT_OPPONENTATTACK1 As Byte = 15
Public Const EFFECT_OPPONENTDEFENSE1 As Byte = 16
Public Const EFFECT_OPPONENTSPEED1 As Byte = 17
Public Const EFFECT_OPPONENTACCURACY1 As Byte = 18
Public Const EFFECT_OPPONENTEVASIVENESS1 As Byte = 19
Public Const EFFECT_RESET_ALL_STAGES As Byte = 20
Public Const EFFECT_TWOORTHREEROUNDS As Byte = 21
Public Const EFFECT_MULTIHIT As Byte = 22
Public Const EFFECT_MAY_FLINCH As Byte = 23
Public Const EFFECT_HALPHP_FULLFAIL As Byte = 24
Public Const EFFECT_BADLY_POSION As Byte = 25
Public Const EFFECT_BURN_FREEZE_PARALYZE As Byte = 26
Public Const EFFECT_REST As Byte = 27
Public Const EFFECT_ONE_HIT_KO As Byte = 28
Public Const EFFECT_TWO_TURN As Byte = 29
Public Const EFFECT_DEMAGE_HALF_HP As Byte = 30
Public Const EFFECT_40_DEMAGE As Byte = 31
Public Const EFFECT_MULTITURN As Byte = 32
Public Const EFFECT_CHANGEFORCRITICAL As Byte = 33
Public Const EFFECT_STRIKESTWICE As Byte = 34
Public Const EFFECT_LOSEHALFHPOFDEMAGEIFMISSES As Byte = 35
Public Const EFFECT_INCREASECRITICALHIT As Byte = 36
Public Const EFFECT_RETURNONEOFFOURHP As Byte = 37
Public Const EFFECT_CONFUSEOPPONENT As Byte = 38
Public Const EFFECT_ATTACK2 As Byte = 39
Public Const EFFECT_DEFENSE2 As Byte = 40
Public Const EFFECT_SPEED2 As Byte = 41
Public Const EFFECT_SPATK2 As Byte = 42
Public Const EFFECT_SPDEF2 As Byte = 43
Public Const EFFECT_OPPONENTATTACK2 As Byte = 44
Public Const EFFECT_OPPONENTDEFENSE2 As Byte = 45
Public Const EFFECT_OPPONENTSPEED2 As Byte = 46
Public Const EFFECT_OPPONENTSPATK2 As Byte = 47
Public Const EFFECT_OPPONENTSPDEF2 As Byte = 48
Public Const EFFECT_MAYOPPONENTATTACK1 As Byte = 49
Public Const EFFECT_MAYOPPONENTDEFENSE1 As Byte = 50
Public Const EFFECT_MAYOPPONENTSPEED1 As Byte = 51
Public Const EFFECT_MAYOPPONENTSPATK1 As Byte = 52
Public Const EFFECT_MAYOPPONENTSPDEF1 As Byte = 53
Public Const EFFECT_MAYOPPONENTACCURACY1 As Byte = 54
Public Const EFFECT_MAYCONFUSE As Byte = 55
Public Const EFFECT_MAYPOISION As Byte = 56


Public Const TYPE_NAME_NORMAL As String = "NORMAL"
Public Const TYPE_NAME_FIGHT As String = "FIGHT"
Public Const TYPE_NAME_FLYING As String = "FLYING"
Public Const TYPE_NAME_POISON As String = "POISON"
Public Const TYPE_NAME_GROUND As String = "GROUND"
Public Const TYPE_NAME_ROCK As String = "ROCK"
Public Const TYPE_NAME_BUG As String = "BUG"
Public Const TYPE_NAME_GHOST As String = "GHOST"
Public Const TYPE_NAME_STEEL As String = "STEEL"
Public Const TYPE_NAME_FIRE As String = "FIRE"
Public Const TYPE_NAME_WATER As String = "WATER"
Public Const TYPE_NAME_GRASS As String = "GRASS"


Public Const PVP_MOVE As Long = 1
Public Const PVP_SWITCH As Long = 2

Public Const DIVISION_BRONZE_3 As Long = 1
Public Const DIVISION_BRONZE_2 As Long = 2
Public Const DIVISION_BRONZE_1 As Long = 3
Public Const DIVISION_SILVER_3 As Long = 4
Public Const DIVISION_SILVER_2 As Long = 5
Public Const DIVISION_SILVER_1 As Long = 6
Public Const DIVISION_GOLD_3 As Long = 7
Public Const DIVISION_GOLD_2 As Long = 8
Public Const DIVISION_GOLD_1 As Long = 9
Public Const DIVISION_PLATINUM_3 As Long = 10
Public Const DIVISION_PLATINUM_2 As Long = 11
Public Const DIVISION_PLATINUM_1 As Long = 12
Public Const DIVISION_DIAMOND_3 As Long = 13
Public Const DIVISION_DIAMOND_2 As Long = 14
Public Const DIVISION_DIAMOND_1 As Long = 15

Public Const DIALOG_NPCBATTLE As Byte = 1
Public Const DIALOG_GIVEITEM As Byte = 2
