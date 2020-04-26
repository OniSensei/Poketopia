Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Sub Main()
    Call InitServer
End Sub

Sub InitServer()
On Error Resume Next
    Dim i As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    If LCase$(Dir(App.Path & "\Data\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\Data\items")
    End If

    If LCase$(Dir(App.Path & "\Data\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Data\maps")
    End If

    If LCase$(Dir(App.Path & "\Data\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir(App.Path & "\Data\npcs")
    End If

    If LCase$(Dir(App.Path & "\Data\shops", vbDirectory)) <> "shops" Then
        Call MkDir(App.Path & "\Data\shops")
    End If

    If LCase$(Dir(App.Path & "\Data\spells", vbDirectory)) <> "spells" Then
        Call MkDir(App.Path & "\Data\spells")
    End If

    If LCase$(Dir(App.Path & "\data\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\data\accounts")
    End If

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Port = 7001
        Options.MOTD = "Welcome to Eclipse Origins"
        Options.Website = "http://www.touchofdeathforums.com/smf/"
        SaveOptions
    Else
        LoadOptions
    End If
    
    'Load EXP
    For i = 1 To 100
    PokemonEXP(i) = Val(GetVar(App.Path & "\data\exp.ini", "EXPERIENCE", "EXP" & i))
    Next
    
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    frmServer.lblstatus = "Loaded in " & (time2 - time1) / 1000 & " seconds. Update version code - " & VersionCode
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Sub DestroyServer()
On Error Resume Next
    Dim i As Long
    ServerOnline = False
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Unloading sockets...")

    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next

    End
End Sub

Sub SetStatus(ByVal status As String)
On Error Resume Next
    Call TextAdd(status)
    NewDoEvents
End Sub

Public Sub ClearGameData()
On Error Resume Next
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTiles
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing Resources...")
    Call ClearResources
    Call SetStatus("Clearing Pokemon...")
    Call ClearPokemons
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing animations...")
    Call ClearAnimations
    Call SetStatus("Clearing moves...")
    Call ClearMoves
End Sub

Private Sub LoadGameData()
On Error Resume Next
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading Pokemon...")
    Call LoadPokemon
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading animations...")
    Call LoadAnimations
    Call SetStatus("Loading moves...")
    Call LoadMove
    Call SetStatus("Loading natures...")
    Call LoadNature
    Call SetStatus("Loading types...")
    Call LoadType
    Call SetStatus("Loading news...")
    Call LoadNews
End Sub

Public Sub TextAdd(msg As String)
On Error Resume Next
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean
On Error Resume Next
    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function

Sub LoadNews()
On Error Resume Next
NewsHTML = ReadText("newsRTF.txt")
End Sub

Function GetNews() As String
GetNews = ReadText("newsRTF.txt")
End Function
