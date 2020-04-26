Attribute VB_Name = "modGlobals"
Option Explicit
'char creation sprite builder
Public spriteIndex As Long
Public spriteGender As String
Public hairIndex As Long
Public hairColor As Long

' pokemon
Public BattleType As Byte
Public enemyPokemon As PokemonBattleEnemyRec
Public BattlePokemon As Long
Public PokemonInstance(1 To 6) As PokemonInstanceRec
Public StorageInstance(1 To 250) As PokemonInstanceRec
Public storagenum As Long
Public selectedpoke As Long
Public BattleRound As Long
' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean

Public isInStorage As Boolean
Public isInBank As Boolean

Public isWaitingForPCScan As Boolean
Public isChatVisible As Boolean
Public MapMusic As String
Public PlayingMapMusic As String

Public AdminOnly As Boolean

Public DialogImage(1 To 100) As Long
Public Dialog(1 To 100) As String
Public Dialogs As Long
Public CurrentDialog As Long
Public IsDialogTrigger(1 To 100) As Boolean

Public AutoCloseBattle As Boolean

Public CurrentNpcX As Long
Public CurrentNpcY As Long
Public TradeLocked As Long
Public TradeName As String
Public WaitingStarter As Long
'
Public MusicIndex As Long
' GUI
Public DragInvSlotNum As Integer
Public InvX As Long
Public InvY As Long
Public EqX As Long
Public EqY As Long
Public SpellX As Long
Public SpellY As Long
Public InvItemFrame(1 To MAX_INV) As Byte ' Used for animated items
Public LastItemDesc As Long ' Stores the last item we showed in desc
Public LastSpellDesc As Long ' Stores the last spell we showed in desc
Public tmpDropItem As Long
Public InShop As Long ' is the player in a shop?
Public ShopAction As Byte ' stores the current shop action

' Player variables
Public MyIndex As Long ' Index of actual player
Public PlayerInv(1 To MAX_INV) As PlayerInvRec   ' Inventory
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Byte
Public InventoryItemSelected As Integer
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Debug mode
Public DEBUG_MODE As Boolean

' Game text buffer
Public MyText As String
Public TextSendTo As String

' TCP variables
Public PlayerBuffer As String

' Used for parsing String packets
Public SEP_CHAR As String * 1
Public END_CHAR As String * 1

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Text variables
Public TexthDC As Long
Public GameFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Used for dragging Picture Boxes
Public SOffsetX As Integer
Public SOffsetY As Integer

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BFPS As Boolean
Public BLoc As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long

' Text vars
Public vbQuote As String

' Mouse cursor tile location
Public CurX As Integer
Public CurY As Integer

' Game editors
Public Editor As Byte
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public SpawnNpcNum As Byte
Public SpawnNpcDir As Byte
Public EditorShop As Long
Public EditorGymBlockNum As Long
Public EditorGymBlockDir As Long
' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key open editor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map Resources
Public ResourceEditorNum As Long

' Maximum classes
Public Max_Classes As Byte
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte

' fps lock
Public FPS_Lock As Boolean

' Editor edited items array
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public NPC_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_NPCS) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Pokemon_Changed(1 To MAX_POKEMONS) As Boolean
Public Move_changed(1 To MAX_MOVES) As Boolean
Public StarterChoosed As Long
' New char
Public newCharSprite As Long
Public newCharClass As Long
Public AreOverWorldsLoaded As Boolean

Public ChatFocus As Boolean

Public isBattleBlocked As Boolean

Public GroundUnvisible As Boolean
Public MaskUnvisible As Boolean
Public Mask2Unvisible As Boolean
Public FringeUnvisible As Boolean
Public Fringe2Unvisible As Boolean

Public ReceivingTime As Long
Public FlashLight As Boolean
Public TPRemoveSlot As Long

Public InPVP As Long

'Intro
Public DrawOak As Boolean
Public DrawEevee As Boolean
Public InIntro As Boolean

'Battle
Public inBattle As Boolean

