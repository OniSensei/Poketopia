Attribute VB_Name = "modScripts"
'Scripting
'READ / WRITE
Function ReadText(ByVal FileName As String) As String
On Error Resume Next
If FileExist(FileName) Then
Open App.Path & "\" & FileName For Input As #1
ReadText = Input$(LOF(1), #1)
Close #1
Else
'WriteText App.Path & "\" & FileName, ""
ReadText = ""
End If
End Function
Sub WriteText(ByVal FileName As String, ByVal Text As String)
On Error Resume Next
Open FileName For Output As #1
Print #1, Text
Close #1
End Sub



Sub DoNpcScript(ByVal index As Long, ByVal script As String)
On Error Resume Next
If Trim$(script) = "" Then Exit Sub
Dim i As Long
Dim arr() As String
Dim command As String
Dim header As String
Dim value As String
Dim value_splitted() As String
Dim value1 As String
Dim value2 As String
Dim x As String
Dim lines() As String
Dim instrchk
Dim a As Long
lines = Split(Trim$(script), vbLf)


For i = 0 To UBound(lines)
instrchk = InStr(1, lines(i), "@")
If instrchk <> 0 Then
arr = Split(lines(i), "@")
command = arr(0)
If UBound(arr) > 0 Then
header = arr(1)
If UBound(arr) > 1 Then
value = arr(2)
value_splitted = Split(value, "=")
value1 = value_splitted(0)
value2 = value_splitted(1)
End If
End If
Select Case command
'-------------------------------------
Case "dialog"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
SendDialog index, value2, Val(arr(3))
End If
'-------------------------------------
Case "putvar"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header, value2)
End If
'-------------------------------------
Case "warp"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
PlayerWarp index, Val(value2), Val(arr(3)), Val(arr(4))
End If

Case "give_item"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
GiveItem index, Val(value2), Val(arr(3))
End If

Case "give_pokemon"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
GivePokemon index, Val(value2), Val(arr(3))
End If

Case "Global"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
GlobalMsg value2, White
End If
'-------------------------------------
Case "stop"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
Call PutVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header, value2)
Exit Sub 'This will stop all
End If
'------------------------------------
Case "shop"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
Dim buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteInteger SOpenShop ' send packet opening the shop
buffer.WriteLong value2
SendDataTo index, buffer.ToArray()
Set buffer = Nothing
TempPlayer(index).InShop = value2 ' stops movement and the like
Exit Sub 'This will stop all
End If
'-------------------------------------
Case "checkpokemon"
Dim foundPoke As Boolean
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
For a = 1 To 6
If Trim$(Pokemon(player(index).PokemonInstance(a).PokemonNumber).Name) = value2 Then
If player(index).PokemonInstance(a).level = Val(arr(3)) Then
foundPoke = True
Exit For
End If
End If
Next
If foundPoke = False Then Exit Sub
End If
'-------------------------------------
Case "takepokemon"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
For a = 1 To 6
If Trim$(Pokemon(player(index).PokemonInstance(a).PokemonNumber).Name) = value2 Then
If player(index).PokemonInstance(a).level = Val(arr(3)) Then
Call TakePlayerPokemon(index, a)
Exit For
End If
End If
Next
End If
'-------------------------------------
Case "givepokemon"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
Dim pLvlCheck As Long
If x = value1 Then
Call GivePokemon(index, PokeNameToNum(value2), Val(arr(3)), 0, 1, 0)
End If
Case "checkpowerlevel"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
For a = 1 To 6
If player(index).PokemonInstance(a).PokemonNumber > 0 Then
pLvlCheck = pLvlCheck + player(index).PokemonInstance(a).level
End If
Next
If pLvlCheck >= Val(value2) Then
Else
Exit Sub
End If
End If

Case "checkitem"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
Dim itSlot As Long
If x = value1 Then
For a = 1 To MAX_INV
If GetPlayerInvItemNum(index, a) = Val(value2) Then
If GetPlayerInvItemValue(index, a) >= Val(arr(3)) Then
itSlot = YES
End If
End If
Next
If itSlot <> YES Then Exit Sub
End If

Case "takeitem"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
Dim itSlot1 As Long
If x = value1 Then
itSlot1 = GetItemSlot(index, Val(value2))
If itSlot1 > 0 Then
If GetPlayerInvItemValue(index, itSlot1) >= Val(arr(3)) Then
Call TakeItem(index, GetPlayerInvItemNum(index, itSlot1), arr(3))
Else
Exit Sub
End If
End If
End If

Case "checkOtherNpc"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
Dim xy As String
If x = value1 Then
xy = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", value2)
If xy = Trim$(arr(3)) Then
Else
Exit Sub
End If
End If
Case "loadnpcbattle"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
LoadNPCBattle index, Val(value2)
End If


Case "startnpcbattle"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
StartNPCBattle index, value2
End If
'-------------------------------------
Case "checkgym"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
If player(index).Bedages(Val(value2)) <> YES Then
Exit Sub
End If
End If
'-------------------------------------
Case "dialognpcbattle"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
SendDialog index, value2, Val(arr(3)), YES
'StartNPCBattle index, value2
TempPlayer(index).hasDialogTrigger = True
TempPlayer(index).dialogTriggerData1 = DIALOG_NPCBATTLE
TempPlayer(index).dialogTriggerData2 = 0
TempPlayer(index).dialogTriggerData3 = Trim$(arr(4))
End If
'-------------------------------------
Case "dialogitem"
x = GetVar(App.Path & "\Data\alive\" & Trim$(player(index).Name) & ".ini", "ALIVE", header)
If x = value1 Then
SendDialog index, value2, Val(arr(3)), YES
'StartNPCBattle index, value2
TempPlayer(index).hasDialogTrigger = True
TempPlayer(index).dialogTriggerData1 = DIALOG_GIVEITEM
TempPlayer(index).dialogTriggerData2 = Val(arr(4))
TempPlayer(index).dialogTriggerData3 = arr(5)
End If
'-------------------------------------


End Select
End If
Next
End Sub










Sub CustomScript(ByVal index As Long, ByVal script As Long)
On Error Resume Next
Dim n As String
Select Case script
Case 1
'This is travel
SendTravel index
Case 2
'FISHING

End Select
End Sub

Public Function PokeNameToNum(ByVal Name As String) As Long
On Error Resume Next
If Name = vbNullString Or Name = "" Then Exit Function
PokeNameToNum = Val(GetVar(App.Path & "\Data\Pokemon Data\Names_Nums.ini", "DATA", Name))
End Function
Sub ItemCustomScript(ByVal index As Long, ByVal itemNum As Long)
On Error Resume Next
Dim x As Long
Select Case itemNum
Case 15
CustomPoke index, 133, 1, NO
Case 17
x = Rand(1, 3)
Select Case x
Case 1
CustomPoke index, 152, 1, NO
Case 2
CustomPoke index, 155, 1, NO
Case 3
CustomPoke index, 158, 1, NO
End Select
Case 18
CustomPoke index, 255, 1, NO
Case 19
CustomPoke index, 258, 1, NO
Case 21
CustomPoke index, 252, 1, NO
Case 23
SetPlayerMembership index, 7
Case 24
SetPlayerMembership index, 30
Case 25
SetPlayerMembership index, 180
Case 26
SetPlayerMembership index, 365

Case 41
x = Rand(1, 3)
Select Case x
Case 1
CustomPoke index, 495, 1, NO
Case 2
CustomPoke index, 498, 1, NO
Case 3
CustomPoke index, 501, 1, NO
End Select

Case 45
CustomPoke index, 1, 1, NO
Case 46
CustomPoke index, 4, 1, NO
Case 47
CustomPoke index, 7, 1, NO

Case 43
If GetHoneyPokes(index) > 0 Then
initHoneyBattle (index)
TakeItem index, itemNum, 1
PlayerMsg index, "Pokemon got attracted by honey!", BrightGreen
Else
PlayerMsg index, "There is no pokemon attracted by honey here.", BrightGreen
End If

Case 52
CustomPoke index, 495, 1, NO
Case 53
CustomPoke index, 498, 1, NO
Case 54
CustomPoke index, 501, 1, NO

Case 55
CustomPoke index, 535, 1, NO
Case 59
CustomPoke index, 408, 1, NO
Case 70
If DoesPlayerHaveEgg(index) = False Then
AddEgg index
PlayerMsg index, "You equipped egg! You can check how close are you to its hatching in egg menu!", Yellow
TakeItem index, itemNum, 1
Else
PlayerMsg index, "You already have and egg equipped!", Yellow
End If
Case 72
If DoesPlayerHaveBike(index) Then
   PlayerMsg index, "You already own a bike!", Yellow
Else
  Call PutVar(App.Path & "\Data\alive\" & GetPlayerName(index) & ".ini", "Other", "Bike", "YES")
  PlayerMsg index, "You now own a bike! To use it press B!", Yellow
  TakeItem index, itemNum, 1
End If



End Select
End Sub



