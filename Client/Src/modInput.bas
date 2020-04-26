Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub CheckKeys()
'Normal
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
'WASD
    'If GetAsyncKeyState(vbKeyW) >= 0 Then DirUp = False
    'If GetAsyncKeyState(vbKeyS) >= 0 Then DirDown = False
    'If GetAsyncKeyState(vbKeyA) >= 0 Then DirLeft = False
    'If GetAsyncKeyState(vbKeyD) >= 0 Then DirRight = False
'other
    If GetAsyncKeyState(VK_SPACE) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
End Sub

Public Sub CheckInputKeys()

    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyReturn) < 0 Then
        CheckMapGetItem
    End If

    If GetKeyState(vbKeySpace) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If

    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirUp = False
    End If

    'Move Right
    If GetKeyState(vbKeyRight) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        Exit Sub
    Else
        DirRight = False
    End If

    'Move down
    If GetKeyState(vbKeyDown) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirDown = False
    End If

    'Move left
    If GetKeyState(vbKeyLeft) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        Exit Sub
    Else
        DirLeft = False
    End If
    
    If ChatFocus = False Then
    
    'Move Up (W)
        If GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right (D)
        If GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down (S)
        If GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left (A)
        If GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        
       
        
        
       End If

End Sub

Public Sub HandleKeypresses(ByVal KeyAscii As Integer)
On Error Resume Next
    Dim ChatText As String
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim Command() As String
    Dim Buffer As clsBuffer
    ChatText = Trim$(MyText)

    If LenB(ChatText) = 0 Then Exit Sub
    MyText = LCase$(ChatText)

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn And isChatVisible = True Then
        ChatFocus = False
        If Trim$(TextSendTo) = "" Then
        Else
        SendRequest 0, 0, ChatText, "PRIVATEM", Trim$(TextSendTo)
        frmChat.txtMyChat.text = vbNullString
        Exit Sub
        End If
        ' Broadcast message
        If Left$(ChatText, 1) = "'" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call BroadcastMsg(ChatText)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If
        
        ' CLAN MSG
        If Left$(ChatText, 1) = "." Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
             SendRequest 0, 0, ChatText, "CLANMSG"
                'Call BroadcastMsg(ChatText)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)

            If Len(ChatText) > 0 Then
                Call EmoteMsg(ChatText)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Player message
        If Left$(ChatText, 1) = "!" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)

                If Mid$(ChatText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If

            Next

            ChatText = Mid$(ChatText, i, Len(ChatText) - 1)

            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
            Case "/intro"
            If inBattle Then Exit Sub
            frmMainGame.Visible = False
            frmChat.Visible = False
            frmIntro.Show
            
            
              Case "/moodhappy"
              SendMood 0
              AddText "Mood set to happy!", BrightBlue
              Case "/heal"
              If GetPlayerAccess(MyIndex) >= 3 Then
              Call SendRequest(0, 0, "", "HEAL")
              End If
              Case "/moodsad"
              AddText "Mood set to sad!", BrightBlue
                SendMood 1
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /help, /info, /who, /fps, /stats, /trade, /party, /join, /leave, /resetui", HelpColor)
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CPlayerInfoRequest
                    Buffer.WriteLong TCP_CODE
                    Buffer.WriteString Command(1)
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CGetStats
                    Buffer.WriteLong TCP_CODE
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                Case "/party"

                    ' Make sure they are actually sending something
                    If UBound(Command) < 1 Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If

                    Call SendPartyRequest(Command(1))
                    ' Join party
                Case "/join"
                    SendJoinParty
                    ' Leave party
                Case "/leave"
                    SendLeaveParty
                    
                Case "/options"
                'frmMainGame.pnlOptions.Visible = True
                'frmMainGame.Check1 = Options.PlayMusic
                'frmMainGame.Check2 = Options.repeatmusic
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("""msghere = Global Admin Message", HelpColor)
                    Call AddText("=msghere = Private Admin Message", HelpColor)
                    Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell, /debug", HelpColor)
                    ' Kicking a player
                    
                    
                 Case "/mapmusic"
                 If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                  GoTo Continue
                  End If
                    
                    If UBound(Command) < 1 Then
                    GoTo Continue
                    End If
                    
                    SetMapMusic Command(1)
                Case "/kick"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/mute"

                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /mute (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /mute (name)", AlertColor
                        GoTo Continue
                    End If

                    SendRequest 0, 0, Command(1), "MUTE"
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/makeclan"
                If UBound(Command) < 1 Then
                    AddText "Usage: /makecrew (name)", AlertColor
                        GoTo Continue
                End If
                SendRequest 0, 0, Trim$(Command(1)), "MAKECREW"
                
                Case "/findpokemon"
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                If IsNumeric(Command(1)) Then
                Call SendRequest(Val(Command(1)), 0, "", "FINDPOKEMON")
                End If
                End If
                
                Case "/visible"
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                Call SendRequest(0, 0, "", "VISIBLE")
                End If
                
                Case "/findpokemonbyname"
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                Call SendRequest(0, 0, Command(1), "FINDPOKEMONBYNAME")
                End If
                Case "/loc"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    BLoc = Not BLoc
                    ' Map Editor
                 Case "/givepokemon"
                If GetPlayerAccess(MyIndex) >= 4 Then
                SendRequest Command(1), 0, "", "GPOKE"
                End If
                Case "/giveitem"
                If GetPlayerAccess(MyIndex) >= 4 Then
                SendRequest Command(1), 0, "", "GITEM"
                End If
                Case "/emote"
                 If GetPlayerAccess(MyIndex) >= 4 Then
                SendRequest 0, 0, Mid$(ChatText, 7, Len(ChatText) - 1), "EMOTE"
                End If
                Case "/mapeditor"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Map report
                Case "/mapreport"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendData CMapReport & END_CHAR
                    ' Respawn request
                Case "/respawn"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                    ' Check the ban list
                Case "/banlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendBanList
                    ' Banning a player
                Case "/ban"

                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditItem
                Case "/editpokemon"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditPokemon
                ' Editing animation request
                Case "/editanimation"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditNpc
                Case "/editresource"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditResource
                    ' Editing shop request
                    
                   
                    
                Case "/editshop"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"

                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"

                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            frmChat.txtMyChat.text = vbNullString
            Exit Sub
        End If

        ' Say message
        If Len(ChatText) > 0 Then
            Call SayMsg(ChatText)
        End If

        MyText = vbNullString
        frmChat.txtMyChat.text = vbNullString
        Exit Sub
    End If

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then

            ' Make sure the character is on standard English keyboard
            If KeyAscii >= 32 Then ' Asc(" ")
                If KeyAscii <= 126 Then ' Asc("~")
                    MyText = MyText & ChrW$(KeyAscii)
                End If
            End If
        End If
    End If

End Sub
