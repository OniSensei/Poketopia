Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public DX7 As New DirectX7  ' Master Object, early binding

Public Sub Main()
   Dim i As Long
    'frmSendGetData.Visible = True
    frmMainGame.lblSGInfo.Visible = True
    Call SetStatus("Loading...")
    GettingMap = True
    vbQuote = ChrW$(34) ' "
    load frmMainGame
    
    ' Update the form with the game's name before it's loaded
    frmMainGame.Caption = GAME_NAME
    
    Call SetStatus("Loading Options...")
    
    ' load options
    LoadOptions
    'natures
    Call SetStatus("Loading Natures...")
    LoadNature
    ' randomize rnd's seed
    Randomize
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call InitMessages
    Call SetStatus("Initializing DirectX...")
    
    ' DX7 Master Object is already created, early binding
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckOverWorld
    '

    
    Call LoadWeather
    ' DirectDraw Surface memory management setting
    DDSD_Temp.lFlags = DDSD_CAPS
    DDSD_Temp.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    
    ' temp set music/sound vars
    Music_On = True
    Sound_On = True
    
    ' load music/sound engine
    InitSound
    InitMusic
    
    ' check if we have main-menu music
    If Len(Trim$(Options.music)) > 0 Then PlayMidi Trim$(Options.music), 1
    
    ' Reset values
    Ping = -1
    
    'Load frmMainMenu ' this line also initalizes directX
    load frmMenu
    frmMenu.Visible = True
    
    ' hide all pics
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picNewChar.Visible = False
    frmMenu.picRegister.Visible = False
    
    ' hide the load form
    'frmSendGetData.Visible = False
    'frmMainGame.lblSGInfo.Visible = False

    '

   
    
End Sub

Public Sub MenuState(ByVal State As Long)
    'frmSendGetData.Visible = True
     frmMainGame.lblSGInfo.Visible = True
    Select Case State
        Case MENU_STATE_ADDCHAR
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMainGame.picNewChar.Visible = False
            frmMainGame.Picture1.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending character addition data...")

                If frmMainGame.optMale.Value Then
                    Call SendAddChar(frmMainGame.txtCName, SEX_MALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite, StarterChoosed, hairColor, hairIndex)
                Else
                    Call SendAddChar(frmMainGame.txtCName, SEX_FEMALE, frmMenu.cmbClass.ListIndex + 1, newCharSprite, StarterChoosed, hairColor, hairIndex)
                End If
            End If
            
        Case MENU_STATE_NEWACCOUNT
            
            frmMainGame.picNewChar.Visible = False
            frmMainGame.Picture1.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMainGame.txtRUser.text, frmMainGame.txtRPass.text)
            End If

        Case MENU_STATE_LOGIN
            frmMenu.Visible = False
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picNewChar.Visible = False
            frmMenu.picRegister.Visible = False

            If ConnectToServer(1) Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMenu.txtLUser.text, frmMenu.txtLPass.text)
                Exit Sub
            End If
    End Select

    If frmMainGame.lblSGInfo.Visible Then
        If Not IsConnected Then
            frmMenu.Visible = True
            frmMenu.picCredits.Visible = False
            frmMenu.picLogin.Visible = False
            frmMenu.picNewChar.Visible = False
            frmMenu.picRegister.Visible = False
            'frmSendGetData.Visible = False
             frmMainGame.lblSGInfo.Visible = False
            Call MsgBox("Server is offline!", vbOKOnly, GAME_NAME)
        End If
    End If

End Sub

Sub GameInit()
    Unload frmMenu
    
    ' Set font
    Call SetFont(FONT_NAME, FONT_SIZE)
    'frmSendGetData.Visible = False
     frmMainGame.lblSGInfo.Visible = False
    frmMainGame.Show
    frmChat.Show
    ' Set the focus
    frmMainGame.picScreen.Visible = True
    frmMainGame.picLogin.Visible = False
    frmMainGame.picHover.Visible = False
    frmMainGame.tmrmenu.Enabled = True
    
    ' Blt inv
    'BltInventory
    frmBag.LoadInv
    frmMainGame.BagLoadInv
    
    If frmPokemons.Visible = True Then frmPokemons.LoadInv
    If frmMainGame.picPokemons.Visible = True Then frmMainGame.RosterLoadInv
    ' set values for amdin panel
    frmAdmin.scrlAItem.Max = MAX_ITEMS
    frmAdmin.scrlAItem.Value = 1
    
    
    
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
    
    ReceivingTime = 0
    'stop the song playing
    StopMidi
End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    InGame = False
    Call DestroyTCP
    
    'destroy objects in reverse order
    Call DestroyDirectDraw

    ' destory DirectX7 master object
    If Not DX7 Is Nothing Then
        Set DX7 = Nothing
    End If

    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
    Dim frm As Form

    For Each frm In VB.Forms

        Unload frm
    Next

End Sub

Public Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
    frmMainGame.lblSGInfo.Caption = frmMainGame.lblSGInfo.Caption & vbNewLine & Caption
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As textbox, Msg As String, NewLine As Boolean)

    If NewLine Then
        Txt.text = Txt.text + Msg + vbCrLf
    Else
        Txt.text = Txt.text + Msg
    End If

    Txt.SelStart = Len(Txt.text) - 1
End Sub

Public Sub SetFocusOnChat()

    On Error Resume Next 'prevent RTE5, no way to handle error

    'frmChat.txtMyChat.SetFocus
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Sub MovePicture(PB As picturebox, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GlobalX As Integer
    Dim GlobalY As Integer
    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + x - SOffsetX
        PB.Top = GlobalY + y - SOffsetY
    End If

End Sub

Public Function isLoginLegal(ByVal Username As String, ByVal Password As String) As Boolean

    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(Password)) >= 3 Then
            isLoginLegal = True
        End If
    End If

End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
    Dim i As Long

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, GAME_NAME)
            Exit Function
        End If

    Next

    isStringLegal = True
End Function


Sub LoadWeather()
Dim i As Long
Dim a As Long
For i = 1 To MAX_MAPS
Weather(i).PicName = GetVar(App.Path & "\Data Files\maps\Weather\" & i & ".ini", "Weather", "type")
Weather(i).Pics = Val(GetVar(App.Path & "\Data Files\maps\Weather\" & i & ".ini", "Weather", "pics"))
Weather(i).speed = Val(GetVar(App.Path & "\Data Files\maps\Weather\" & i & ".ini", "Weather", "speed"))
If Weather(i).Pics > 0 Then
For a = 1 To Weather(i).Pics
Weather(i).pics_Y(a) = 0
Weather(i).pics_x(a) = Rand(10, 700)
Next
End If
Next
End Sub

Function GetMyProcess() As String
Dim process As Object
Dim a As String

For Each process In GetObject("winmgmts:").ExecQuery("Select * from Win32_Process")
    a = a & vbNewLine & process.Caption
Next
GetMyProcess = a
End Function

