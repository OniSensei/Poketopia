Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long

Public Const ERROR_SUCCESS As Long = 0
Public Const BINDF_GETNEWESTVERSION As Long = &H10
Public Const INTERNET_FLAG_RELOAD As Long = &H80000000

'version
Public Const VersionCode As Long = 1133
Public Const TCP_CODE As Long = 432001
Public Const BATTLE_NO As Long = 5
' Inventory constants
Public Const InvTop As Byte = 8
Public Const InvLeft As Byte = 8
Public Const InvOffsetY As Byte = 4
Public Const InvOffsetX As Byte = 4
Public Const InvColumns As Byte = 5

' spells constants
Public Const SpellTop As Byte = 8
Public Const SpellLeft As Byte = 8
Public Const SpellOffsetY As Byte = 4
Public Const SpellOffsetX As Byte = 4
Public Const SpellColumns As Byte = 5

' shop constants
Public Const ShopTop As Byte = 8
Public Const ShopLeft As Byte = 8
Public Const ShopOffsetY As Byte = 4
Public Const ShopOffsetX As Byte = 4
Public Const ShopColumns As Byte = 5

' Character consts
Public Const EqTop As Byte = 200
Public Const EqLeft As Byte = 8
Public Const EqOffsetX As Byte = 15
Public Const EqColumns As Byte = 4
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

' Font variables
'Public Const FONT_NAME As String = "Verdana Bold"
Public Const FONT_NAME As String = "Eurostar"
Public Const FONT_SIZE As Byte = 14

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\Data Files\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\Data Files\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\Data Files\graphics\"
Public Const GFX_EXT As String = ".bmp"
Public Const POKE_PATH As String = "\Data Files\graphics\pokemonsprites\"
' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_W As Long = &H87
Public Const VK_A As Long = &H65
Public Const VK_S As Long = &H83
Public Const VK_D As Long = &H68
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11
Public Const VK_SPACE As Long = &H20

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Number of tiles in width in tilesets
Public Const TILESHEET_WIDTH As Integer = 15 ' * PIC_X pixels

' Speed moving vars
Public Const WALK_SPEED As Byte = 2
Public Const RUN_SPEED As Byte = 4
Public Const BIKE_SPEED As Byte = 8
' Tile size constants
Public Const PIC_X As Integer = 32
Public Const PIC_Y As Integer = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Integer = 32
Public Const SIZE_Y As Integer = 32

' ********************************************************
' * The values below must match with the server's values *
' ********************************************************
'Public Const ERROR_SUCCESS As Long = 0
'Public Const BINDF_GETNEWESTVERSION As Long = &H10
'Public Const INTERNET_FLAG_RELOAD As Long = &H80000000

' General constants
Public Const GAME_NAME As String = "Poketopia"
Public Const MAX_PLAYERS As Long = 500
Public Const MAX_ITEMS As Byte = 255
Public Const MAX_NPCS As Byte = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Byte = 35
Public Const MAX_MAP_ITEMS As Byte = 255
Public Const MAX_MAP_NPCS As Byte = 30
Public Const MAX_MAP_POKEMONS As Long = 30
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
' Website
Public Const GAME_WEBSITE As String = "http://www.peocommunity.boards.net"

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
Public Const MAX_MAPX As Byte = 21
Public Const MAX_MAPY As Byte = 19
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

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

' Constants for player movement: Tiles per Second
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
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_POKEMON As Byte = 7

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2
Public Const HalfX As Integer = ((MAX_MAPX + 1) \ 2) * PIC_X
Public Const HalfY As Integer = ((MAX_MAPY + 1) \ 2) * PIC_Y
Public Const ScreenX As Integer = (MAX_MAPX + 1) * PIC_X
Public Const ScreenY As Integer = (MAX_MAPY + 1) * PIC_Y

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

'TYPES
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

'GymBlock Direction
Public Const GYMBLOCK_DIRECTION_UPDOWN As Long = 5
Public Const GYMBLOCK_DIRECTION_LEFT As Long = 7
Public Const GYMBLOCK_DIRECTION_RIGHT As Long = 6
'Gyms
Public Const GYM_DEFEATED As Long = 1
Public Const GYM_UNDEFEATED As Long = 0


Public Const MENU_TRAINERCARD As Byte = 1
Public Const MENU_POKEDEX As Byte = 2
Public Const MENU_OPTIONS As Byte = 3
Public Const MENU_ROSTER As Byte = 4
Public Const MENU_BAG As Byte = 5
Public Const MENU_PROFILE As Byte = 6
Public Const MENU_BANK As Byte = 7
Public Const MENU_EVOLVE As Byte = 8
Public Const MENU_LEARNMOVE As Byte = 9
Public Const MENU_TRADE As Byte = 10
Public Const MENU_TRAVEL As Byte = 11
Public Const MENU_SHOP As Byte = 12
Public Const MENU_TPREMOVE As Byte = 13
Public Const MENU_CREW As Byte = 14
Public Const MENU_EGG As Byte = 15

Public Const GDI_IMAGE_NIGHT As Byte = 0
Public Const GDI_IMAGE_OAK As Byte = 1
Public Const GDI_IMAGE_EEVEE As Byte = 2
