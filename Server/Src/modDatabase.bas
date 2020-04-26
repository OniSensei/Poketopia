Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
On Error Resume Next
    Dim FileName As String
    Dim F As Integer

    If ServerLog Then
        FileName = App.Path & "\data\logs\" & FN

        If Not FileExist(FileName, True) Then
            F = FreeFile
            Open FileName For Output As #F
            Close #F
        End If

        F = FreeFile
        Open FileName For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(file As String, header As String, Var As String) As String
On Error Resume Next
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(file As String, header As String, Var As String, value As String)
On Error Resume Next
    Call WritePrivateProfileString$(header, Var, value, file)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
On Error Resume Next
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(FileName)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    On Error Resume Next
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", str(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    
End Sub

Public Sub LoadOptions()
    On Error Resume Next
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
On Error Resume Next
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
     FileName = App.Path & "\data\banlist.ini"
        
    ' Check if file exists
    IP = GetPlayerIP(BanPlayerIndex)
    Call PutVar(FileName, "DATA", IP, "YES")
    Call PutVar(FileName, "DATA", GetPlayerName(BanPlayerIndex), "YES")

  

    ' Cut off last portion of ip
    

   
    Call GlobalMsg("[BAN!] " & GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", Cyan)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long, Optional ByVal reason As String = "")
On Error Resume Next
    Dim FileName As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    FileName = App.Path & "\data\banlist.ini"
        
    ' Check if file exists
    IP = GetPlayerIP(BanPlayerIndex)
    Call PutVar(FileName, "DATA", IP, "YES")
    Call PutVar(FileName, "DATA", GetPlayerName(BanPlayerIndex), "YES")
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, reason + " You have been banned by " & "The Server" & "!")
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
On Error Resume Next
    Dim FileName As String
    FileName = "data\accounts\" & Trim(Name) & ".bin"

    If FileExist(FileName) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
On Error Resume Next
    Dim FileName As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Integer

    If AccountExist(Name) Then
        FileName = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        Get #nFileNum, NAME_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
On Error Resume Next
    Dim i As Long
    
    ClearPlayer Index
    
    player(Index).Login = Name
    player(Index).Password = Password

    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
On Error Resume Next
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal Index As Long) As Boolean
On Error Resume Next
    If LenB(Trim$(player(Index).Name)) > 0 Then
        CharExist = True
    End If

End Function

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal Sprite As Long, Starter As Long, hairC As Long, hairI As Long)
On Error Resume Next
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(player(Index).Name)) = 0 Then
        
        spritecheck = False
        
        player(Index).Name = Name
        player(Index).Sex = Sex
        player(Index).Class = ClassNum
        
        If player(Index).Sex = SEX_MALE Then
            For n = 0 To UBound(Class(ClassNum).MaleSprite)
                If Class(ClassNum).MaleSprite(n) = Sprite Then
                    spritecheck = True
                End If
            Next
        Else
            For n = 0 To UBound(Class(ClassNum).FemaleSprite)
                If Class(ClassNum).FemaleSprite(n) = Sprite Then
                    spritecheck = True
                End If
            Next
        End If
        
        ' Sprite not valid, simply reset to '1'
        If Not spritecheck Then
            If player(Index).Sex = SEX_MALE Then
                Sprite = Class(ClassNum).MaleSprite(0)
            Else
                Sprite = Class(ClassNum).FemaleSprite(0)
            End If
        End If
        
        player(Index).Sprite = Sprite

        player(Index).level = 1

        For n = 1 To Stats.Stat_Count - 1
            player(Index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        player(Index).Dir = DIR_DOWN
        player(Index).map = START_MAP
        player(Index).x = START_X
        player(Index).y = START_Y
        player(Index).Dir = DIR_DOWN
        player(Index).Vital(Vitals.hp) = GetPlayerMaxVital(Index, Vitals.hp)
        player(Index).Vital(Vitals.mp) = GetPlayerMaxVital(Index, Vitals.mp)
        player(Index).Vital(Vitals.SP) = GetPlayerMaxVital(Index, Vitals.SP)
        player(Index).SX = 14
        player(Index).SY = 6
        player(Index).SMap = 1
         'Pokeball shirt
        GivePokemon Index, Starter, 5, 0, NO, 33
        If Trim$(GetVar(App.Path & "\Data\testeralive\" & Name & ".ini", "Other", "Logged")) = "YES" Then
         SetPlayerInvItemNum Index, 3, 17
        SetPlayerInvItemValue Index, 3, 1
        PlayerMsg Index, "Thank you for being a tester!", Yellow
        End If
        SetPlayerInvItemNum Index, 1, 1 ' give player start money
        SetPlayerInvItemValue Index, 1, 1000 ' set player start money value
        SetPlayerInvItemNum Index, 2, 2 ' give player start pokeballs
        SetPlayerInvItemValue Index, 2, 5 ' set player start pokeballs qty
        If player(Index).Sex = SEX_MALE Then
        Select Case hairC
            Case 1
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 255
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 244
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 233
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 222
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 2
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 254
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 243
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 232
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 221
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 3
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 253
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 242
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 231
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 220
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 4
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 252
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 241
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 230
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 219
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 5
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 251
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 240
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 229
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 218
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 6
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 250
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 239
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 228
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 217
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 7
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 249
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 238
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 227
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 216
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 8
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 248
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 237
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 226
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 215
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 9
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 247
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 236
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 225
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 214
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 10
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 246
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 235
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 224
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 213
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 11
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 245
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 234
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 223
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 212
                        SetPlayerInvItemValue Index, 3, 1
                End Select
        End Select
        Else
        Select Case hairC
            Case 1
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 211
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 200
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 189
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 178
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 2
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 210
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 199
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 188
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 177
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 3
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 209
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 198
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 187
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 176
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 4
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 208
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 197
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 186
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 175
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 5
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 207
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 196
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 185
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 174
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 6
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 206
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 195
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 184
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 173
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 7
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 205
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 194
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 183
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 172
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 8
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 204
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 193
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 182
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 171
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 9
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 203
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 192
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 181
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 170
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 10
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 202
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 191
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 180
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 169
                        SetPlayerInvItemValue Index, 3, 1
                End Select
            Case 11
                Select Case hairI
                    Case 1
                        SetPlayerInvItemNum Index, 3, 201
                        SetPlayerInvItemValue Index, 3, 1
                    Case 2
                        SetPlayerInvItemNum Index, 3, 190
                        SetPlayerInvItemValue Index, 3, 1
                    Case 3
                        SetPlayerInvItemNum Index, 3, 179
                        SetPlayerInvItemValue Index, 3, 1
                    Case 4
                        SetPlayerInvItemNum Index, 3, 168
                        SetPlayerInvItemValue Index, 3, 1
                End Select
        End Select
        End If
        
        SetPlayerInvItemNum Index, 4, 166 ' give player first shirt
        SetPlayerInvItemValue Index, 4, 1 ' set shirt qty
        SetPlayerInvItemNum Index, 5, 167 ' give player first pants
        SetPlayerInvItemValue Index, 5, 1 ' set pants qty
        SetPlayerInvItemNum Index, 6, 165 ' give player first backpack
        SetPlayerInvItemValue Index, 6, 1 ' set backpack qty
        PlayerMsg Index, "You received starting items. Check your bag to equip them.", Yellow
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
On Error Resume Next
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SavePlayer(i)
            If TempPlayer(i).eggExpTemp >= 1000 Then
            SaveEggFromTemp i
            End If
           If TempPlayer(i).eggStepsTemp >= 100 Then
           SaveEggFromTemp i
           End If
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long

    FileName = App.Path & "\data\accounts\" & Trim$(player(Index).Login) & ".bin"
    
    F = FreeFile
    
    Open FileName For Binary As #F
    Put #F, , player(Index)
    Close #F
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    Call ClearPlayer(Index)
    FileName = App.Path & "\data\accounts\" & Trim(Name) & ".bin"
    F = FreeFile
    Open FileName For Binary As #F
    Get #F, , player(Index)
    Close #F
End Sub

Sub ClearPlayer(ByVal Index As Long)
On Error Resume Next
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(Index)), LenB(TempPlayer(Index)))
    Set TempPlayer(Index).buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(player(Index)), LenB(player(Index)))
    player(Index).Login = vbNullString
    player(Index).Password = vbNullString
    player(Index).Name = vbNullString
    player(Index).Class = 1

    frmServer.lvwInfo.ListItems(Index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(Index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
On Error Resume Next
    Dim FileName As String
    Dim file As String
    FileName = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(FileName, True) Then
        file = FreeFile
        Open FileName For Output As file
        Print #file, "[INIT]"
        Print #file, "MaxClasses=" & Max_Classes
        Close file
    End If

End Sub

Sub LoadClasses()
On Error Resume Next
    Dim FileName As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        FileName = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.strength) = Val(GetVar(FileName, "CLASS" & i, "Str"))
        Class(i).Stat(Stats.endurance) = Val(GetVar(FileName, "CLASS" & i, "End"))
        Class(i).Stat(Stats.vitality) = Val(GetVar(FileName, "CLASS" & i, "Vit"))
        Class(i).Stat(Stats.willpower) = Val(GetVar(FileName, "CLASS" & i, "Will"))
        Class(i).Stat(Stats.intelligence) = Val(GetVar(FileName, "CLASS" & i, "Int"))
        Class(i).Stat(Stats.spirit) = Val(GetVar(FileName, "CLASS" & i, "Spir"))
    Next

End Sub





Sub SaveClasses()
On Error Resume Next
    Dim FileName As String
    Dim i As Long
    FileName = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Maleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(FileName, "CLASS" & i, "Str", CStr(Class(i).Stat(Stats.strength)))
        Call PutVar(FileName, "CLASS" & i, "End", CStr(Class(i).Stat(Stats.endurance)))
        Call PutVar(FileName, "CLASS" & i, "Vit", CStr(Class(i).Stat(Stats.vitality)))
        Call PutVar(FileName, "CLASS" & i, "Will", CStr(Class(i).Stat(Stats.willpower)))
        Call PutVar(FileName, "CLASS" & i, "Int", CStr(Class(i).Stat(Stats.intelligence)))
        Call PutVar(FileName, "CLASS" & i, "Spr", CStr(Class(i).Stat(Stats.spirit)))
    Next

End Sub

Function CheckClasses() As Boolean
On Error Resume Next
    Dim FileName As String
    FileName = App.Path & "\data\classes.ini"

    If Not FileExist(FileName, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
On Error Resume Next
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
    Next

End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next

End Sub

Sub SaveItem(ByVal itemNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F  As Long
    FileName = App.Path & "\data\items\item" & itemNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , item(itemNum)
    Close #F
End Sub

Sub LoadItems()
On Error Resume Next
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , item(i)
        Close #F
    Next

End Sub

Sub CheckItems()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(item(Index)), LenB(item(Index)))
    item(Index).Name = vbNullString
End Sub

Sub ClearItems()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next

End Sub

Sub SaveShop(ByVal ShopNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\shops\shop" & ShopNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
On Error Resume Next
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        FileName = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

End Sub

Sub ClearShop(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
End Sub

Sub ClearShops()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal spellnum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\spells\spells" & spellnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Spell(spellnum)
    Close #F
End Sub

Sub SaveSpells()
On Error Resume Next
    Dim i As Long
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next

End Sub

Sub LoadSpells()
On Error Resume Next
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        FileName = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).LevelReq = 1 'Needs to be 1 for the spell editor
End Sub

Sub ClearSpells()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

End Sub

' **********
' ** NPCs **
' **********
Sub SaveNpcs()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , NPC(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , NPC(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
End Sub

Sub ClearNpcs()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        FileName = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_RESOURCES

        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If

    Next

End Sub

Sub ClearResource(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
End Sub

Sub ClearResources()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** Pokemon **
' **********
Sub SavePokemons()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        Call SavePokemon(i)
    Next

End Sub

Sub SavePokemon(ByVal pokemonnum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\Pokemon\Pokemon" & pokemonnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Pokemon(pokemonnum)
    Close #F
End Sub

Sub LoadPokemon()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckPokemon

    For i = 1 To MAX_POKEMONS
        FileName = App.Path & "\data\Pokemon\Pokemon" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Pokemon(i)
        Close #F
    Next

End Sub

Sub CheckPokemon()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_POKEMONS

        If Not FileExist("\Data\Pokemon\Pokemon" & i & ".dat") Then
            Call SavePokemon(i)
        End If

    Next

End Sub

Sub ClearPokemon(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(Pokemon(Index)), LenB(Pokemon(Index)))
    Pokemon(Index).Name = vbNullString
End Sub

Sub ClearPokemons()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_POKEMONS
        Call ClearPokemon(i)
    Next
End Sub
' **********
' ** moves *
' **********
Sub SaveMoves()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MOVES
        Call SaveMove(i)
    Next

End Sub

Sub SaveMove(ByVal moveNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\moves\Move" & moveNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , PokemonMove(moveNum)
    Close #F
End Sub

Sub LoadMove()
On Error Resume Next
     Dim FileName As String
     Dim Fn2 As String
    Dim i As Integer
    Dim a As Integer
    Dim fn3 As String
    Dim categ As String
    FileName = App.Path & "\Data\Moves.ini"
    Fn2 = App.Path & "\Data\MovesData\"
    fn3 = App.Path & "\Data\MovesINI\"
   ' For i = 1 To MAX_MOVES
        'PokemonMove(i).Name = Trim$(GetVar(FileName, "MOVE" & i, "Name"))
    'Next
    
    For a = 1 To 621
       PokemonMove(a).Name = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Name")
       PokemonMove(a).Description = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Description")
       categ = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Category")
       If categ = "Status" Then
       PokemonMove(a).Category = "Other Damage"
       End If
       If categ = "Physical" Then
       PokemonMove(a).Category = "Physical Damage"
       End If
       If categ = "Special" Then
       PokemonMove(a).Category = "Special Damage"
       End If
       PokemonMove(a).Type = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Type")
       PokemonMove(a).pp = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "PP"))
       PokemonMove(a).power = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "Power"))
       PokemonMove(a).accuracy = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "Accuracy"))
       PokemonMove(a).Generation = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "Gen"))
       PokemonMove(a).doesDamageIfMiss = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "doesDamageIfMiss"))
       PokemonMove(a).missDamageModifier = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "missDamageModifier"))
       PokemonMove(a).isFlinching = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isFlinching"))
       PokemonMove(a).flinchChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "flinchChanses"))
       PokemonMove(a).isCharging = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isCharging"))
       PokemonMove(a).canBeHitOnCharging = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "canBeHitOnCharging"))
       PokemonMove(a).chargeFirst = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "chargeFirst"))
       PokemonMove(a).chargeTurns = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "chargeTurns"))
       PokemonMove(a).critical_hit_ration = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "critical-hit-ratio"))
       PokemonMove(a).isMultiTurn = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isMultiTurn"))
       PokemonMove(a).multiTurnLowerLimit = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "multiTurnLowerLimit"))
       PokemonMove(a).multiTurnUpperLimit = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "multiTurnUpperLimit"))
       PokemonMove(a).isMultiTurn = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isMultiHit"))
       PokemonMove(a).multiHitLowerLimit = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "multiHitLowerLimit"))
       PokemonMove(a).multiHitUpperLimit = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "multiHitUpperLimit"))
       PokemonMove(a).isRecoil = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isRecoil"))
       PokemonMove(a).isHealing = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isHealing"))
       PokemonMove(a).priority = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "priority"))
       PokemonMove(a).recoilModifier = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "recoilModifier"))
       PokemonMove(a).isAttackerStatChanging = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isAttackerStatChanging"))
       PokemonMove(a).attackerStatChangeModifier = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "attackerStatChangeModifier"))
       PokemonMove(a).attackerStatChangeIndex = GetVar(Fn2 & a + 1 & ".ini", "DATA", "attackerStatChangeIndex")
       PokemonMove(a).attackerStatChangeChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "attackerStatChangeChances"))
       PokemonMove(a).isOpponentStatChanging = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isOpponentStatChanging"))
       PokemonMove(a).opponentStatChangeModifier = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentStatChangeModifier"))
       PokemonMove(a).opponentStatChangeIndex = GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentStatChangeIndex")
       PokemonMove(a).opponentStatChangeChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentStatChangeChances"))
       PokemonMove(a).isOpponentNonVolatileStatusInducing = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isOpponentNonVolatileStatusInducing"))
       PokemonMove(a).opponentNonVolatileStatusInducingChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentNonVolatileStatusInducingChances"))
       PokemonMove(a).nonVolatileStatusType = GetVar(Fn2 & a + 1 & ".ini", "DATA", "nonVolatileStatusType")
       PokemonMove(a).isOpponentStatResetting = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isOpponentStatResetting"))
       PokemonMove(a).opponentStatResettingChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentStatResettingChances"))
       PokemonMove(a).isAttackerStatResetting = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isAttackerStatResetting"))
       PokemonMove(a).attackerStatResettingChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "attackerStatResettingChances"))
       PokemonMove(a).isHpRestoring = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isHpRestoring"))
       PokemonMove(a).HpRestoringChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "HpRestoringChances"))
       PokemonMove(a).hpRestoreModifier = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "hpRestoreModifier"))
       PokemonMove(a).isOpponentVolatileStatusIndiucing = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isOpponentVolatileStatusIndiucing"))
       PokemonMove(a).opponentVolatileStatusInducingChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "opponentVolatileStatusInducingChances"))
       PokemonMove(a).OpponentVolatileStatusType = GetVar(Fn2 & a + 1 & ".ini", "DATA", "OpponentVolatileStatusType")
       PokemonMove(a).isAttackerVolatileStatusInducing = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "isAttackerVolatileStatusInducing"))
       PokemonMove(a).AttackerVolatileStatusType = GetVar(Fn2 & a + 1 & ".ini", "DATA", "AttackerVolatileStatusType")
       PokemonMove(a).attackerVolatileStatusInducingChances = Val(GetVar(Fn2 & a + 1 & ".ini", "DATA", "attackerVolatileStatusInducingChances"))
       PokemonMove(a).InteralType1 = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Interal Type 1")
       PokemonMove(a).InteralType2 = GetVar(Fn2 & a + 1 & ".ini", "DATA", "Interal Type 2")
       MoveTypes(a) = MoveTextToType(Trim$(PokemonMove(a).Type))
    Next
End Sub

Sub CheckMove()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MOVES

        If Not FileExist("\Data\moves\Move" & i & ".dat") Then
            Call SaveMove(i)
        End If

    Next

End Sub

Sub ClearMove(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(PokemonMove(Index)), LenB(PokemonMove(Index)))
    PokemonMove(Index).Name = vbNullString
End Sub

Sub ClearMoves()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MOVES
        Call ClearMove(i)
    Next
End Sub


' **********
' ** animations **
' **********
Sub SaveAnimations()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        FileName = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal Index As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).Name = vbNullString
End Sub

Sub ClearAnimations()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal mapnum As Long)
On Error Resume Next
    Dim FileName As String
    Dim F As Long
    Dim x As Long
    Dim y As Long
    FileName = App.Path & "\data\maps\map" & mapnum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , map(mapnum).Name
    Put #F, , map(mapnum).Revision
    Put #F, , map(mapnum).Moral
    Put #F, , map(mapnum).Tileset
    Put #F, , map(mapnum).Up
    Put #F, , map(mapnum).Down
    Put #F, , map(mapnum).Left
    Put #F, , map(mapnum).Right
    Put #F, , map(mapnum).Music
    Put #F, , map(mapnum).BootMap
    Put #F, , map(mapnum).BootX
    Put #F, , map(mapnum).BootY
    Put #F, , map(mapnum).MaxX
    Put #F, , map(mapnum).MaxY

    For x = 0 To map(mapnum).MaxX
        For y = 0 To map(mapnum).MaxY
            Put #F, , map(mapnum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , map(mapnum).NPC(x)
    Next

    For x = 1 To MAX_MAP_POKEMONS
    Put #F, , map(mapnum).Pokemon(x).PokemonNumber
    Put #F, , map(mapnum).Pokemon(x).LevelFrom
    Put #F, , map(mapnum).Pokemon(x).LevelTo
    Put #F, , map(mapnum).Pokemon(x).Custom
    Put #F, , map(mapnum).Pokemon(x).atk
    Put #F, , map(mapnum).Pokemon(x).def
    Put #F, , map(mapnum).Pokemon(x).spatk
    Put #F, , map(mapnum).Pokemon(x).spdef
    Put #F, , map(mapnum).Pokemon(x).spd
    Put #F, , map(mapnum).Pokemon(x).hp
    Put #F, , map(mapnum).Pokemon(x).Chance
    Next

    Close #F
    NewDoEvents
End Sub

Sub SaveMaps()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next

End Sub

Sub LoadMaps()
On Error Resume Next
    Dim FileName As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        FileName = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , map(i).Name
        Get #F, , map(i).Revision
        Get #F, , map(i).Moral
        Get #F, , map(i).Tileset
        Get #F, , map(i).Up
        Get #F, , map(i).Down
        Get #F, , map(i).Left
        Get #F, , map(i).Right
        Get #F, , map(i).Music
        Get #F, , map(i).BootMap
        Get #F, , map(i).BootX
        Get #F, , map(i).BootY
        Get #F, , map(i).MaxX
        Get #F, , map(i).MaxY
        ' have to set the tile()
        ReDim map(i).Tile(0 To map(i).MaxX, 0 To map(i).MaxY)

        For x = 0 To map(i).MaxX
            For y = 0 To map(i).MaxY
                Get #F, , map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , map(i).NPC(x)
            MapNpc(i).NPC(x).Num = map(i).NPC(x)
        Next
        
        For x = 1 To MAX_MAP_POKEMONS
        Get #F, , map(i).Pokemon(x).PokemonNumber
        Get #F, , map(i).Pokemon(x).LevelFrom
        Get #F, , map(i).Pokemon(x).LevelTo
        Get #F, , map(i).Pokemon(x).Custom
        Get #F, , map(i).Pokemon(x).atk
        Get #F, , map(i).Pokemon(x).def
        Get #F, , map(i).Pokemon(x).spatk
        Get #F, , map(i).Pokemon(x).spdef
        Get #F, , map(i).Pokemon(x).spd
        Get #F, , map(i).Pokemon(x).hp
        Get #F, , map(i).Pokemon(x).Chance
        Next
        
        Close #F
        
        'Dim mapnpcs As Long
        'mapnpcs = 0
        
        'For x = 0 To map(i).MaxX
          'For y = 0 To map(i).MaxX
             'If FileExist("Data\NPCScripts\" & i & "I" & x & "I" & y & ".ini") Then
               
                'mapnpcs = mapnpcs + 1
                'map(i).NPCNames(mapnpcs) = GetVar(App.Path & "\Data\NPCScripts\" & i & "I" & x & "I" & y & ".ini", "DATA", "Name")
                'map(i).NPCScripts(mapnpcs) = ReadText("Data\NPCScripts\" & i & "I" & x & "I" & y & ".txt")
                'map(i).NPCX(mapnpcs) = x
                'map(i).NPCY(mapnpcs) = y
            ' End If
          'Next
        'Next
        
        ClearTempTile i
        CacheResources i
        NewDoEvents
    Next
End Sub






Sub CheckMaps()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, Index)), LenB(MapItem(mapnum, Index)))
End Sub

Sub ClearMapItems()
On Error Resume Next
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal mapnum As Long)
On Error Resume Next
    ReDim MapNpc(mapnum).NPC(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).NPC(Index)), LenB(MapNpc(mapnum).NPC(Index)))
End Sub

Sub ClearMapNpcs()
On Error Resume Next
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub clearmappokemon(ByVal Index As Long, ByVal mapnum As Long)

End Sub

Sub ClearMap(ByVal mapnum As Long)
On Error Resume Next
    Call ZeroMemory(ByVal VarPtr(map(mapnum)), LenB(map(mapnum)))
    map(mapnum).Tileset = 1
    map(mapnum).Name = vbNullString
    map(mapnum).MaxX = MAX_MAPX
    map(mapnum).MaxY = MAX_MAPY
    ReDim map(mapnum).Tile(0 To map(mapnum).MaxX, 0 To map(mapnum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Sub ClearMaps()
On Error Resume Next
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
On Error Resume Next
If ClassNum < 1 Or ClassNum > Max_Classes Then
GetClassName = "CLASS"
Else
GetClassName = Trim$(Class(ClassNum).Name)
End If
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
On Error Resume Next
    Select Case Vital
        Case hp
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.vitality) \ 2) + Class(ClassNum).Stat(Stats.vitality)) * 2
        Case mp
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.intelligence) \ 2) + Class(ClassNum).Stat(Stats.intelligence)) * 2
        Case SP
            GetClassMaxVital = (1 + (Class(ClassNum).Stat(Stats.spirit) \ 2) + Class(ClassNum).Stat(Stats.spirit)) * 2
    End Select

End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
On Error Resume Next
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function


'/////////////NATURES/////////////
Sub LoadNature()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    FileName = App.Path & "\Data\Natures.ini"
    For i = 1 To MAX_NATURES
        nature(i).Name = GetVar(FileName, "NATURE" & i, "Name")
        nature(i).AddHP = Val(GetVar(FileName, "NATURE" & i, "HP"))
        nature(i).AddAtk = Val(GetVar(FileName, "NATURE" & i, "ATK"))
        nature(i).AddDef = Val(GetVar(FileName, "NATURE" & i, "DEF"))
        nature(i).AddSpAtk = Val(GetVar(FileName, "NATURE" & i, "SPATK"))
        nature(i).AddSpDef = Val(GetVar(FileName, "NATURE" & i, "SPDEF"))
        nature(i).AddSpd = Val(GetVar(FileName, "NATURE" & i, "SPEED"))
    Next
End Sub

Sub LoadType()
On Error Resume Next
    Dim FileName As String
    Dim i As Integer
    FileName = App.Path & "\Data\Types.ini"
    For i = 1 To 18
        Types(i).NORMAL = Val(GetVar(FileName, "TYPE" & i, "NORMAL"))
        Types(i).FIGHT = Val(GetVar(FileName, "TYPE" & i, "FIGHT"))
        Types(i).FLYING = Val(GetVar(FileName, "TYPE" & i, "FLYING"))
        Types(i).POISON = Val(GetVar(FileName, "TYPE" & i, "POISON"))
        Types(i).GROUND = Val(GetVar(FileName, "TYPE" & i, "GROUND"))
        Types(i).ROCK = Val(GetVar(FileName, "TYPE" & i, "ROCK"))
        Types(i).BUG = Val(GetVar(FileName, "TYPE" & i, "BUG"))
        Types(i).GHOST = Val(GetVar(FileName, "TYPE" & i, "GHOST"))
        Types(i).STEEL = Val(GetVar(FileName, "TYPE" & i, "STEEL"))
        Types(i).FIRE = Val(GetVar(FileName, "TYPE" & i, "FIRE"))
        Types(i).WATER = Val(GetVar(FileName, "TYPE" & i, "WATER"))
        Types(i).GRASS = Val(GetVar(FileName, "TYPE" & i, "GRASS"))
        Types(i).ELECTRIC = Val(GetVar(FileName, "TYPE" & i, "ELECTRIC"))
        Types(i).PSYCHIC = Val(GetVar(FileName, "TYPE" & i, "PSYCHIC"))
        Types(i).ICE = Val(GetVar(FileName, "TYPE" & i, "ICE"))
        Types(i).DRAGON = Val(GetVar(FileName, "TYPE" & i, "DRAGON"))
        Types(i).DARK = Val(GetVar(FileName, "TYPE" & i, "DARK"))
        Types(i).FAIRY = Val(GetVar(FileName, "TYPE" & i, "FAIRY"))
    Next
End Sub

Public Function MoveTextToType(ByVal Text As String) As Byte
On Error Resume Next
Select Case Text
Case "NONE"
MoveTextToType = TYPE_NONE
Case "NORMAL"
MoveTextToType = TYPE_NORMAL
Case "BUG"
MoveTextToType = TYPE_BUG
Case "DARK"
MoveTextToType = TYPE_DARK
Case "DRAGON"
MoveTextToType = TYPE_DRAGON
Case "ELECTRIC"
MoveTextToType = TYPE_ELECTRIC
Case "FAIRY"
MoveTextToType = TYPE_FAIRY
Case "FIGHTING"
MoveTextToType = TYPE_FIGHTING
Case "FIRE"
MoveTextToType = TYPE_FIRE
Case "FLYING"
MoveTextToType = TYPE_FLYING
Case "GHOST"
MoveTextToType = TYPE_GHOST
Case "GRASS"
MoveTextToType = TYPE_GRASS
Case "GROUND"
MoveTextToType = TYPE_GROUND
Case "ICE"
MoveTextToType = TYPE_ICE
Case "POISON"
MoveTextToType = TYPE_POISON
Case "PSYCHIC"
MoveTextToType = TYPE_PSYCHIC
Case "ROCK"
MoveTextToType = TYPE_ROCK
Case "STEEL"
MoveTextToType = TYPE_STEEL
Case "WATER"
MoveTextToType = TYPE_WATER
End Select
End Function

Public Function IsItemTradeable(ByVal itemNum As Long) As Boolean
On Error Resume Next
Dim iString As String
iString = itemNum
If GetVar(App.Path & "\Data\Tradeable.ini", "DATA", iString) = "NO" Then
IsItemTradeable = False
Else
IsItemTradeable = True
End If
End Function


'CREWS

Function DoesCrewExist(ByVal crew As String) As Boolean
On Error Resume Next
If GetVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Exists") = "YES" Then
DoesCrewExist = True
End If
End Function

Sub MakeCrew(ByVal leader As Long, ByVal crewname As String)
On Error Resume Next
If DoesCrewExist(crewname) = False Then
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Picture", "https://cdn.pixabay.com/photo/2015/04/11/10/08/shield-717505_960_720.png")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Leader", GetPlayerName(leader))
Call PutVar(App.Path & "\Data\alive\" & GetPlayerName(leader) & ".ini", "DATA", "Crew", crewname)
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Exists", "YES")
Dim i As Long
For i = 1 To 50
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Member" & i, "")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Member" & i & "Admin", "NO")
Next
For i = 1 To 30
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Request" & i, "")
Next

GlobalMsg "Clan " & crewname & " was created by " & GetPlayerName(leader) & "!", BrightGreen
End If
End Sub


Sub DeleteCrew(ByVal crewname As String)
On Error Resume Next
If DoesCrewExist(crewname) Then
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Picture", "")
Call PutVar(App.Path & "\Data\alive\" & GetCrewLeaderName(crewname) & ".ini", "DATA", "Crew", "")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Leader", "")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Exists", "NO")
Dim i As Long
For i = 1 To 50
If GetPlayerCrewByName(GetCrewPlayerName(crewname, i)) = crewname Then
Call PutVar(App.Path & "\Data\alive\" & GetCrewPlayerName(crewname, i) & ".ini", "DATA", "Crew", "")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Member" & i, "")
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Member" & i & "Admin", "NO")
End If

Next
For i = 1 To 30
Call PutVar(App.Path & "\Data\crews\" & crewname & ".ini", "DATA", "Request" & i, "")
Next

GlobalMsg "Clan " & crewname & " was deleted!", BrightGreen
End If
End Sub


Sub AddMemberToCrew(ByVal crew As String, ByVal slot As Long, ByVal member As Long)
On Error Resume Next
If DoesCrewExist(crew) = False Then Exit Sub
Dim i As Long
Call PutVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Member" & slot, GetPlayerName(member))
Call PutVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Member" & slot & "Admin", "NO")
Call PutVar(App.Path & "\Data\alive\" & GetPlayerName(member) & ".ini", "DATA", "Crew", crew)
End Sub

Function GetPlayerCrew(ByVal Index As Long)
On Error Resume Next
GetPlayerCrew = GetVar(App.Path & "\Data\alive\" & GetPlayerName(Index) & ".ini", "DATA", "Crew")
End Function


Function GetPlayerCrewByName(ByVal player As String)
On Error Resume Next
GetPlayerCrewByName = GetVar(App.Path & "\Data\alive\" & Trim$(player) & ".ini", "DATA", "Crew")
End Function



Function GetCrewPlayerName(ByVal crew As String, ByVal slot As Long) As String
If DoesCrewExist(crew) = False Then Exit Function
GetCrewPlayerName = GetVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Member" & slot)
End Function
Function GetCrewLeaderName(ByVal crew As String) As String
If DoesCrewExist(crew) = False Then Exit Function
GetCrewLeaderName = GetVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Leader")
End Function
Function GetCrewPicture(ByVal crew As String) As String
If DoesCrewExist(crew) = False Then Exit Function
GetCrewPicture = GetVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Picture")
End Function

Function GetCrewFreeSpot(ByVal crew As String) As Long
Dim i As Long
For i = 1 To 50
If GetCrewPlayerName(crew, i) = "" Then
GetCrewFreeSpot = i
Exit Function
End If
Next
End Function

Sub AddToCrew(ByVal Index As Long, ByVal crew As String)
Dim spot As Long
spot = GetCrewFreeSpot(crew)
If spot < 1 Or spot > 50 Then
PlayerMsg Index, "There is no free spot in the clan!", Yellow
Exit Sub
End If
AddMemberToCrew crew, spot, Index
ClanMsg crew, GetPlayerName(Index) & " joined clan!"
End Sub
Sub ClanMsg(ByVal clan As String, ByVal msg As String)
Dim i As Long
If FindPlayer(GetCrewLeaderName(clan)) > 0 Then
PlayerMsg FindPlayer(GetCrewLeaderName(clan)), "[CLAN] " & msg, BrightGreen
End If
For i = 1 To 50
If Trim$(GetCrewPlayerName(clan, i)) <> vbNullString Then
If FindPlayer(GetCrewPlayerName(clan, i)) > 0 Then
If IsPlaying(FindPlayer(GetCrewPlayerName(clan, i))) Then
PlayerMsg FindPlayer(GetCrewPlayerName(clan, i)), "[CLAN] " & msg, BrightGreen
Else
End If
Else
End If
Else
End If
Next
End Sub

Public Function GetPlayerCrewSpot(ByVal Index As Long, ByVal crew As String) As Long
Dim i As Long
For i = 1 To 50
If GetCrewPlayerName(crew, i) = GetPlayerName(Index) Then
GetPlayerCrewSpot = i
Exit Function
End If
Next

End Function

Sub RemoveMemberFromCrew(ByVal crew As String, ByVal slot As Long)
On Error Resume Next
Dim member As String
If DoesCrewExist(crew) = False Then Exit Sub
Dim i As Long
member = GetCrewPlayerName(crew, slot)
Call PutVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Member" & slot, "")
Call PutVar(App.Path & "\Data\crews\" & crew & ".ini", "DATA", "Member" & slot & "Admin", "NO")
Call PutVar(App.Path & "\Data\alive\" & member & ".ini", "DATA", "Crew", "")
End Sub




Function GetCrewNews(ByVal crew As String) As String
GetCrewNews = ReadText("Data\clanNews\" & crew & ".txt")
End Function


Public Function DoesItemTake(ByVal itemNum As Long) As Boolean
On Error Resume Next
If GetVar(App.Path & "\Data\TakeScriptItems.ini", "DATA", Trim$(itemNum)) = "YES" Then
DoesItemTake = True
End If
End Function
