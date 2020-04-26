Attribute VB_Name = "modPokemon"

'//////////////////BATTLE ///////////////////////////
'//////////////////SYSTEM////////////////////////////

Function May() As Long
On Error Resume Next
Dim i As Long
i = Rand(1, 100)
If i <= 30 Then May = YES

End Function

Function SpawnChance(ByVal OneOf As Long) As Boolean
On Error Resume Next
'n = Int(Rnd * OneOf) + 1
Dim x As Long
Dim y As Long
x = Rand(1, OneOf)
y = Rand(1, OneOf)
'If n = 1 then
If x = y Then
SpawnChance = True
Else
SpawnChance = False
End If
End Function

Function SpawnChanceDecimal(ByVal OneOf As Double) As Boolean
On Error Resume Next
n = Int(Rnd * OneOf) + 1
If n = 1 Then
SpawnChanceDecimal = True
Else
SpawnChanceDecimal = False
End If
End Function



Function DealDemage(ByVal index As Long, ByVal move As Long, MoveType As String, ByVal AttackerATK As Long, ByVal defenderdef As Long, isCritical As Long, attackertype As Byte, attackerType2 As Byte, ByVal attackerlvl As Long, ByVal defendertype As Byte, ByVal defenderType2 As Byte, Optional ByVal isCustomMove As Boolean = False, Optional ByVal custommove As Long = 1, Optional ByVal customPower As Boolean, Optional ByVal customPowerVal As Long, Optional ByVal powerModifier As Long = 0)
On Error Resume Next
Dim power As Long, accuracy As Long, maxpp As Long, number As Long, pp As Long, level As Long, attack As Long, defense As Long, cr As Long
Dim modifier As Double
Dim stab As Double, mtype As Double, typeModifier As Double
If isCustomMove = True Then
power = PokemonMove(custommove).power
accuracy = PokemonMove(custommove).accuracy
Else
power = PokemonMove(move).power
accuracy = PokemonMove(move).accuracy
End If

If customPower = True Then power = customPowerVal
If powerModifier > 0 Then
power = power * (powerModifier / 100)
End If
maxpp = PokemonMove(move).pp
level = attackerlvl

If power < 1 Then
DealDemage = 0
Exit Function
End If
'Do checks
If isCritical = YES Then
cr = 2
Else
cr = 1
End If
If MoveTypes(move) = attackertype Or MoveTypes(move) = attackerType2 Then
stab = 1.5
Else
stab = 1
End If
'STILL NEEDS SOME CODE .......................................................................
Dim x As Double
Dim y As Double
x = GetTypeEffect(MoveTypes(move), defendertype)
y = GetTypeEffect(MoveTypes(move), defenderType2)
typeModifier = x * y

'.............................................................................................
modifier = stab * typeModifier * cr
Dim xa As Double
Dim xb As Double
Dim xc As Double
Dim xd As Double
Dim xe As Double
Dim xf As Double
Select Case defenderdef
Case Is <= 0
DealDemage = 1000
Case Is > 0
xa = 2 * level + 10
xb = xa / 250
xc = AttackerATK / defenderdef
xd = xc * power + 2
'DealDemage = (((2 * level + 10) / 250) * (AttackerATK / defenderdef) * power + 2) * modifier
DealDemage = (((((((level * 2 / 5) + 2) * power * AttackerATK / 50) / (defenderdef * 1)) * 1) + 2) * cr * 1 * ((Rand(255, 217) * 100) / 255) / 100) * stab * x * y * 1
DealDemage = DealDemage * 0.6
'DealDemage = (xb * xc * power + 2) * modifier
End Select
If DealDemage < 1 And typeModifier > 0 Then
DealDemage = 1
End If
End Function

Function IsWildDefeated(ByVal index As Long) As Boolean
On Error Resume Next
If TempPlayer(index).PokemonBattle.PokemonNumber <= 0 Or TempPlayer(index).PokemonBattle.PokemonNumber > MAX_POKEMONS Then
IsWildDefeated = True
Exit Function
End If

If TempPlayer(index).PokemonBattle.hp <= 0 Then
IsWildDefeated = True
Exit Function
Else
IsWildDefeated = False
Exit Function
End If
End Function

Function isPlayerDefeated(ByVal index As Long, ByVal slot As Long) As Boolean
On Error Resume Next
Dim hp As Long
Dim i As Long
Dim haveMoves As Long

If player(index).PokemonInstance(slot).PokemonNumber <= 0 Then
isPlayerDefeated = True
Exit Function
End If

'Check if player have PP
'For i = 1 To 4
'If player(index).PokemonInstance(slot).moves(i).pp >= 1 Then
'haveMoves = haveMoves + 1
'Else
'End If
'Next
'If haveMoves = 0 Then
'isPlayerDefeated = True
'reason = "You dont have any PP to attack!"
'Exit Function
'End If


'Check HP
hp = player(index).PokemonInstance(slot).hp
If hp < 1 Then
isPlayerDefeated = True
player(index).PokemonInstance(slot).hp = 0
'reason = "Your pokemon is fainted!"
Else
isPlayerDefeated = False
End If

End Function

Sub CheckTurn(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & "->" & Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " Overspeed!", Black
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & "->" & Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " Overspeed!", Black
Else
TempPlayer(index).BattleTurn = False
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & "->" & Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " Overspeed!", Black
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & "->" & Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " Overspeed!", Black
End If
End If
End Sub

Sub BattleCommand(ByVal index As Long, ByVal command As Byte, ByVal slot As Long, ByVal move As Long)
On Error Resume Next

If TempPlayer(index).PokemonBattle.PokemonNumber < 1 Or TempPlayer(index).PokemonBattle.hp < 1 Then Exit Sub
SendBattleUpdate index, slot
Dim pc As Long
Dim i As Long, header As String
Dim exp_gained As Long
'Checking
If command = 1 Then
'If TempPlayer(index).PokemonBattle.PokemonNumber <= 0 Or TempPlayer(index).PokemonBattle.PokemonNumber > MAX_POKEMONS Then Exit Sub
''If TempPlayer(index).PokemonBattle.Hp <= 0 Then Exit Sub


If TempPlayer(index).isInPVP = True Then
If TempPlayer(index).PVPHasUsed = True Then Exit Sub
TempPlayer(index).PVPCommandUsed = PVP_MOVE
TempPlayer(index).PVPCommandNum = move
TempPlayer(index).PVPSlot = slot
TempPlayer(index).PVPHasUsed = True
If TempPlayer(FindPlayer(Trim$(TempPlayer(index).PVPEnemy))).PVPHasUsed = True Then
PVPProcessRound index
End If
Exit Sub
End If


If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
CheckTurn index, slot
If Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(move).number).Name) = "Quick Attack" Or Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(move).number).Name) = "Shadow Sneak" Or Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(move).number).Name) = "Mach Punch" Or Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(move).number).Name) = "Ice Shard" Then
PlayerAttackWild index, move, slot
If WildDefeated(index, slot) Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
Else
Select Case TempPlayer(index).BattleTurn
Case True
PlayerAttackWild index, move, slot
If WildDefeated(index, slot) Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
Case False
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
PlayerAttackWild index, move, slot
If WildDefeated(index, slot) Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
End Select
End If
End If

If command = 2 Then

If TempPlayer(index).isInPVP = True Then
If TempPlayer(index).PVPHasUsed = True Then Exit Sub

TempPlayer(index).PVPCommandUsed = PVP_SWITCH
TempPlayer(index).PVPCommandNum = slot
PlayerMsg index, slot, Yellow
TempPlayer(index).PVPHasUsed = True
If TempPlayer(FindPlayer(Trim$(TempPlayer(index).PVPEnemy))).PVPHasUsed = True Then
PVPProcessRound index
End If
Exit Sub
End If

If TempPlayer(index).WaitingForSwitch = YES Then
SetAsLeader index, slot
SendPlayerPokemon index
SendPlayerData index
SendBattleUpdate index, 1, YES
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
TempPlayer(index).WaitingForSwitch = NO
Exit Sub
Else
SendBattleUpdate index, slot
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
End If
SendBattleUpdate index, slot, YES
SendPlayerPokemon index
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
Exit Sub
End If

If command = 3 Then
If TempPlayer(index).inNPCBattle = True Then Exit Sub
If CheckItem(index, move, slot) = True Then
SendBattleUpdate index, slot
If TempPlayer(index).PokemonBattle.PokemonNumber > 0 Then
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
End If
End If
SendBattleUpdate index, slot, YES
SendPlayerPokemon index
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
Exit Sub
End If

If command = 4 Then 'This is struggle
'If TempPlayer(index).PokemonBattle.PokemonNumber <= 0 Or TempPlayer(index).PokemonBattle.PokemonNumber > MAX_POKEMONS Then Exit Sub
'If TempPlayer(index).PokemonBattle.Hp <= 0 Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
CheckTurn index, slot
Select Case TempPlayer(index).BattleTurn
Case True
PlayerAttackWild index, move, slot, True
If WildDefeated(index, slot) Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
Case False
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
PlayerAttackWild index, move, slot, True
If WildDefeated(index, slot) Then Exit Sub
If PlayerDefeated(index, slot) Then Exit Sub
CheckStatusTurnWild (index)
CheckStatusTurnMine index, slot
End Select
End If


If command = 5 Then 'FLEE
If TempPlayer(index).inNPCBattle = True Then Exit Sub
'Try to flee
If player(index).PokemonInstance(slot).hp > 0 And player(index).PokemonInstance(slot).PokemonNumber > 0 Then
If Runned(index, slot) Then

TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
Call SendBattleInfo(index, pc, 3, 0)
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slot, YES
Else
If CanWildAttack(index, slot) Then
WildAttackPlayer index, slot
If PlayerDefeated(index, slot) Then Exit Sub
If WildDefeated(index, slot) Then Exit Sub
End If
End If
End If
End If


SendPlayerPokemon index
SendBattleUpdate index, slot
End Sub

Sub CheckStatusTurnWild(ByVal index As Long)
On Error Resume Next
If TempPlayer(index).PokemonBattle.status > STATUS_NOTHING Then
TempPlayer(index).PokemonBattle.statusturn = TempPlayer(index).PokemonBattle.statusturn + 1
If TempPlayer(index).PokemonBattle.turnsneed = TempPlayer(index).PokemonBattle.statusturn Then
TempPlayer(index).PokemonBattle.statusturn = 0
TempPlayer(index).PokemonBattle.turnsneed = 0
TempPlayer(index).PokemonBattle.status = STATUS_NOTHING
End If
End If
End Sub

Sub CheckStatusTurnMine(ByVal index As Long, ByVal slt As Long)
On Error Resume Next
If player(index).PokemonInstance(slt).status > STATUS_NOTHING Then
player(index).PokemonInstance(slt).statusstun = player(index).PokemonInstance(slt).statusstun + 1
If player(index).PokemonInstance(slt).turnsneed = player(index).PokemonInstance(slt).statusstun Then
player(index).PokemonInstance(slt).statusstun = 0
player(index).PokemonInstance(slt).turnsneed = 0
player(index).PokemonInstance(slt).status = STATUS_NOTHING
End If
End If

End Sub

Sub CheckWildDefeated(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
Dim pc As Long
If IsWildDefeated(index) = True Then

pc = Rand(1, 3)
If isPlayerMember(index) Then
pc = pc * 1.5
End If
exp_gained = GetExpWildBattle(index, slot)
'GiveItem index, 1, pc
Call GiveEXPtoPlayer(index, slot, exp_gained)
'Close the battle
Call SendBattleInfo(index, pc, YES, exp_gained)
'Call GiveItem(index, 1, pc)
ResetBattlePokemon (index)
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0

SendPlayerPokemon index
SendBattleUpdate index, slot
'SendSound index, "Victory.wav"
Else
End If
End Sub

Sub CheckPlayerDefeated(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
If isPlayerDefeated(index, slot) = True Then
Dim playerPC As Long
playerPC = GetPlayerInvItemValue(index, GetItemSlot(index, 1))
Dim takePC As Long
takePC = playerPC * 0.05
If takePC > 0 Then
Call TakeItem(index, 1, takePC)
End If
Call SendBattleInfo(index, 0, BATTLE_NO, 0)
ResetBattlePokemon (index)
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
SendPlayerPokemon index
SendBattleUpdate index, slot
'Spawn player
'PlayerMsg index, "", BrightRed
Call SpawnPlayer(index)
Call HealPokemons(index)
Else
End If
End Sub

Function WildDefeated(ByVal index As Long, ByVal slot As Long) As Boolean
On Error Resume Next
Dim exp_gained As Long
Dim i As Long
If IsWildDefeated(index) = True Then
If TempPlayer(index).inNPCBattle = True Then
For i = 1 To TempPlayer(index).NPCBattlePokesAvailable
If TempPlayer(index).NPCBattlePokemons(i).PokemonNumber > 0 And TempPlayer(index).NPCBattlePokemons(i).hp > 0 And i <> TempPlayer(index).NPCBattleSelectedPoke Then
'Switch poke
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).NPCBattlePokemons(TempPlayer(index).NPCBattleSelectedPoke).PokemonNumber).Name) & " can't battle any more!"
NPCSwitchPoke index, i, slot
SendPlayerPokemon index
SendBattleUpdate index, slot
WildDefeated = True
Exit Function
End If
Next
CheckNpcWinData index
TempPlayer(index).inNPCBattle = False
TempPlayer(index).NPCBattle = 0

End If
exp_gained = GetExpWildBattle(index, slot)
Dim pc As Long
pc = Rand(1, 3)
If isPlayerMember(index) Then
pc = pc * 1.5
End If
'GiveItem index, 1, pc
Call GiveEXPtoPlayer(index, slot, exp_gained)
'Close the battle
Call GiveItem(index, 1, pc, NO)

WildDrop index
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
Call SendBattleInfo(index, pc, YES, exp_gained)
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slot
'SendSound index, "Victory.wav"
WildDefeated = True
Else
WildDefeated = False
End If
End Function

Function PlayerDefeated(ByVal index As Long, ByVal slot As Long) As Boolean
On Error Resume Next
Dim i As Long
Dim n As Long

If WildDefeated(index, slot) Then
PlayerDefeated = False
Exit Function
End If
If isPlayerDefeated(index, slot) = True Then
For i = 1 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 And player(index).PokemonInstance(i).hp > 0 Then
n = n + 1
End If
Next
If n > 0 Then
TempPlayer(index).WaitingForSwitch = YES

SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " can't battle any more!"
PlayerDefeated = True
SendBattleUpdate index, slot
SendPlayerPokemon index
SendOpenSwitch index
Exit Function
End If

If TempPlayer(index).inNPCBattle = True Then TempPlayer(index).inNPCBattle = False


TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
Dim playerPC As Long
playerPC = GetPlayerInvItemValue(index, GetItemSlot(index, 1))
Dim takePC As Long
takePC = playerPC * 0.05
If takePC > 0 Then
Call TakeItem(index, 1, takePC)
End If
Call SendBattleInfo(index, 0, BATTLE_NO, 0)
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slot
'Spawn player
'PlayerMsg index, reason, BrightRed
Call SpawnPlayer(index)
Call HealPokemons(index)
PlayerDefeated = True
Else
PlayerDefeated = False
End If
End Function

Function CanWildAttack(ByVal index As Long, slot As Long) As Boolean
On Error Resume Next
Dim n As Long
Dim i As Long
If TempPlayer(index).inNPCBattle = True Then
If TempPlayer(index).PokemonBattle.status = STATUS_SLEEPING Then

For i = 1 To TempPlayer(index).NPCBattlePokesAvailable
If TempPlayer(index).NPCBattlePokemons(i).PokemonNumber > 0 And TempPlayer(index).NPCBattlePokemons(i).hp > 0 And i <> TempPlayer(index).NPCBattleSelectedPoke Then
'Switch poke
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).NPCBattlePokemons(TempPlayer(index).NPCBattleSelectedPoke).PokemonNumber).Name) & " - Switched! (Reason: Fell asleep!)"
NPCSwitchPoke index, i, slot
SendPlayerPokemon index
SendBattleUpdate index, slot
Exit Function
End If
Next
End If
End If

Select Case TempPlayer(index).PokemonBattle.status
Case 0
CanWildAttack = True
Case STATUS_NOTHING
CanWildAttack = True
Case STATUS_SLEEPING
CanWildAttack = False
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is still sleeping!", BrightRed
Case STATUS_BURNED
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.hp - (TempPlayer(index).PokemonBattle.MaxHp / 8)
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " burned himself!", BrightRed
n = TempPlayer(index).PokemonBattle.MaxHp / 8
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " lost " & n & " HP!"
If WildDefeated(index, slot) Then
Exit Function
Else
CanWildAttack = True
End If
Case STATUS_PARALIZED
CanWildAttack = False
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is paralized! " & Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " can't move!", BrightRed
Case STATUS_POISONED
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.hp - (TempPlayer(index).PokemonBattle.MaxHp / 8)
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is still poisoned!", BrightRed
n = TempPlayer(index).PokemonBattle.MaxHp / 8
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " lost " & n & " HP!"
If WildDefeated(index, slot) Then
Exit Function
Else
CanWildAttack = True
End If
Case STATUS_FREEZED
CanWildAttack = False
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is frozen.", BrightRed
Case STATUS_CONFUSED
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is confused.", BrightRed
n = TempPlayer(index).PokemonBattle.MaxHp / 8
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.hp - n
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " hurt itself.", BrightRed
Dim hpStr As String
hpStr = n
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " lost " & hpStr & "HP", BrightRed
CanWildAttack = True
Case STATUS_FLINCHED
CanWildAttack = False
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is flinching.", BrightRed
Case STATUS_BADLYPOISONED
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.hp - (TempPlayer(index).PokemonBattle.MaxHp / 6)
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is badly poisoned!", BrightRed
n = TempPlayer(index).PokemonBattle.MaxHp / 6
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " lost " & n & " HP!"
If WildDefeated(index, slot) Then
Exit Function
Else
CanWildAttack = True
End If
End Select
End Function


Sub PlayerAttackWild(ByVal index As Long, ByVal move As Long, ByVal slot As Long, Optional ByVal struggle As Boolean = False)
On Error Resume Next
'Do checks
Dim n As Long
If TempPlayer(index).PokemonBattle.PokemonNumber <= 0 Or TempPlayer(index).PokemonBattle.PokemonNumber > MAX_POKEMONS Then Exit Sub
If TempPlayer(index).PokemonBattle.hp <= 0 Then Exit Sub
If player(index).PokemonInstance(slot).status > STATUS_NOTHING And player(index).PokemonInstance(slot).turnsneed > 0 Then
Select Case player(index).PokemonInstance(slot).status
Case STATUS_NOTHING
Case STATUS_SLEEPING
If Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(move).number).Name) = "Sleep Talk" Then
Else
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is still sleeping!", BrightRed
Exit Sub
End If
Case STATUS_BURNED
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).hp - (player(index).PokemonInstance(slot).MaxHp / 8)
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " burned himself.", BrightRed
n = player(index).PokemonInstance(slot).MaxHp / 8
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " lost " & n & " HP!", BrightRed
If PlayerDefeated(index, slot) = True Then
Exit Sub
End If
Case STATUS_PARALIZED
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is paralized!", BrightRed
Exit Sub
Case STATUS_POISONED
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).hp - (player(index).PokemonInstance(slot).MaxHp / 8)
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is poisoned.", BrightRed
n = player(index).PokemonInstance(slot).MaxHp / 8
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " lost " & n & " HP!", BrightRed
If PlayerDefeated(index, slot) = True Then
Exit Sub
End If
Case STATUS_FREEZED
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is frozen.", BrightRed
Exit Sub
Case STATUS_CONFUSED
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is confused.", BrightRed
n = DealDemage(index, 144, "NONE", player(index).PokemonInstance(slot).batk, player(index).PokemonInstance(slot).bdef, NO, TYPE_NONE, TYPE_NONE, player(index).PokemonInstance(slot).level, "NONE", "NONE", False, 1, True, 40, 0)
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).hp - n
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " hurt itself.", BrightRed
Dim hpStr As String
hpStr = n
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " lost " & hpStr & "HP", BrightRed
Case STATUS_FLINCHED
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is flinching.", BrightRed
Exit Sub
Case STATUS_BADLYPOISONED
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).hp - (player(index).PokemonInstance(slot).MaxHp / 6)
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " is badly poisoned.", BrightRed
n = player(index).PokemonInstance(slot).MaxHp / 6
SendBattleMessage index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " lost " & n & " HP!", BrightRed
If PlayerDefeated(index, slot) = True Then
Exit Sub
End If
End Select
End If


If struggle = True Then
UseMove index, True, 135, slot, move
Exit Sub
End If
If player(index).PokemonInstance(slot).moves(move).pp < 1 Then Exit Sub
UseMove index, True, player(index).PokemonInstance(slot).moves(move).number, slot, move
player(index).PokemonInstance(slot).moves(move).pp = player(index).PokemonInstance(slot).moves(move).pp - 1
End Sub

Sub WildAttackPlayer(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
Dim i As Long
Dim wildAI As Long
If TempPlayer(index).PokemonBattle.PokemonNumber > 0 Then
If TempPlayer(index).inNPCBattle = True Then
wildAI = GetNPCAI(index, slot)

If wildAI > 0 Then
UseMove index, False, wildAI, slot
Else
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is powerless!", Black
End If
Else
If GetVar(App.Path & "\Data\MapScript.ini", GetPlayerMap(index), "CanPokesAttack") <> "NO" Then
wildAI = GetWildAI(index, slot)
If wildAI > 0 Then
UseMove index, False, wildAI, slot
Else
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is powerless!", Black
End If
Else
If TempPlayer(index).PokemonBattle.level > 3 Then
wildAI = GetWildAI(index, slot)
If wildAI > 0 Then
UseMove index, False, wildAI, slot
Else
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is powerless!", Black
End If
Else
SendBattleMessage index, Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & " is powerless!", Black
End If
End If
End If
End If
End Sub

Sub GiveEXPtoPlayer(ByVal index As Long, ByVal slot As Long, ByVal EXP As Long)
On Error Resume Next
player(index).PokemonInstance(slot).EXP = player(index).PokemonInstance(slot).EXP + EXP
TempPlayer(index).eggExpTemp = TempPlayer(index).eggExpTemp + EXP

If player(index).PokemonInstance(slot).level < 100 Then
Call LevelUp(index, slot)
End If
End Sub

Function GetExpWildBattle(ByVal index As Long, ByVal slot As Long) As Long
On Error Resume Next
Dim xp As Long
Dim b As Long
Dim l As Long
Dim Lp As Long

b = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).BaseEXP
l = TempPlayer(index).PokemonBattle.level
Lp = player(index).PokemonInstance(slot).level

Dim tlt As Long
Dim llpt As Long
tlt = (2 * l + 10) * (2 * l + 10) * ((2 * l + 10) / 2)
llpt = (l + Lp + 10) * (l + Lp + 10) * ((l + Lp + 10) / 2)
xp = (b * l) / 7
GetExpWildBattle = xp * 0.55
If isPlayerMember(index) Then
GetExpWildBattle = GetExpWildBattle * 1.2
End If
'event only
If EXP35 = True Then
GetExpWildBattle = GetExpWildBattle * 1.35
End If
End Function

Sub LevelUp(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
'Do Until player(index).PokemonInstance(slot).EXP < PokemonEXP(player(index).PokemonInstance(slot).level + 1)
If player(index).PokemonInstance(slot).level < 100 Then
If player(index).PokemonInstance(slot).EXP >= PokemonEXP(player(index).PokemonInstance(slot).level + 1) Then
player(index).PokemonInstance(slot).EXP = player(index).PokemonInstance(slot).EXP - PokemonEXP(player(index).PokemonInstance(slot).level + 1)
player(index).PokemonInstance(slot).level = player(index).PokemonInstance(slot).level + 1
player(index).PokemonInstance(slot).TP = player(index).PokemonInstance(slot).TP + 3
Call PlayerMsg(index, "Your " & Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " has leveled up!", Yellow)
Call PlayerMsg(index, "Your pokemon has gained 3TP!", Yellow)
Call SendActionMsg(GetPlayerMap(index), "Level up!", BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32)
If (player(index).PokemonInstance(slot).level Mod 10) = 0 Then
NatureBonus index, slot
End If
Call CheckForMoveLearn(index, slot)
End If
Else
'Exit Do
End If
'Loop

End Sub

Sub CheckForMoveLearn(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
Dim pokeNum As Long
pokeNum = player(index).PokemonInstance(slot).PokemonNumber
Dim i As Long
Dim a As Long
Dim n As Long
For i = 1 To 30
n = 0
If Pokemon(pokeNum).movesLV(i) = player(index).PokemonInstance(slot).level Then
If pokeNum = player(index).PokemonInstance(slot).PokemonNumber Then
PlayerMsg index, "[SYSTEM] " & Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " can now learn " & Trim$(PokemonMove(Pokemon(pokeNum).moves(i)).Name), BrightRed
'LEARN MOVE
For a = 1 To 4
If Pokemon(pokeNum).moves(i) = player(index).PokemonInstance(slot).moves(a).number Then
n = n + 1
End If
Next
If n = 0 Then
TempPlayer(index).LearnMoveNumber = Pokemon(pokeNum).moves(i)
TempPlayer(index).LearnMovePokemon = slot
TempPlayer(index).LearnMovePokemonName = Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name)
SendLearnMove index, slot, Pokemon(pokeNum).moves(i)
Exit For
Else

End If
End If
End If
Next
End Sub

Function CalculateStat(ByVal base As Long, ByVal Stat As Long) As Long
On Error Resume Next
Select Case Stat
Case STAT_ATK
CalculateStat = base
Case STAT_DEF
CalculateStat = base
Case STAT_SPATK
CalculateStat = base
Case STAT_SPDEF
CalculateStat = base
Case STAT_SPEED
CalculateStat = base
Case STAT_HP
CalculateStat = base
End Select

End Function


Function Status_Rounds() As Long
On Error Resume Next
'60% to be 1 or 2 rounds
If SpawnChanceDecimal(1.625) Then ' 60% of change to be 1 or 2 rounds
Status_Rounds = Rand(1, 2)
Else
If SpawnChance(5) Then '20% of chance to be 3 or 4 round
Status_Rounds = Rand(3, 4)
Else
If SpawnChance(10) Then '10% of chance to be 5 rounds
Status_Rounds = 5
Else
Status_Rounds = Rand(1, 2) ' If not then 1 or 2
End If
End If

End If

End Function








Sub UseMove(ByVal index As Long, OnOpponent As Boolean, ByVal move As Long, ByVal slt As Long, Optional mslot As Long = 1)
On Error Resume Next
 Dim critical As Long, pp As Long, atype As Byte, atype2 As Byte, deftype1 As Byte, deftype2 As Byte, mtype As String, alvl As Long, aname As String, defname As String, atk As Long, def As Long
Dim n As Long
Dim x As Long
Dim i As Long
Dim hp(1 To 2) As Long
Dim status(1 To 2) As Long
Dim patk(1 To 2) As Long
Dim pdef(1 To 2) As Long
Dim spatk(1 To 2) As Long
Dim spdef(1 To 2) As Long
Dim spd(1 To 2) As Long
Dim MaxHp(1 To 2) As Long
Dim statusturn(1 To 2) As Long
Dim turnsneed(1 To 2) As Long
Dim used As Boolean
Dim miss As Boolean
Dim isCritical As Boolean
mtype = PokemonMove(move).Type
critical = NO
Select Case OnOpponent
Case True
'
patk(1) = player(index).PokemonInstance(slt).batk
pdef(1) = player(index).PokemonInstance(slt).bdef
spatk(1) = player(index).PokemonInstance(slt).bspatk
spdef(1) = player(index).PokemonInstance(slt).bspdef
spd(1) = player(index).PokemonInstance(slt).bspd
hp(1) = player(index).PokemonInstance(slt).hp
MaxHp(1) = player(index).PokemonInstance(slt).MaxHp
status(1) = player(index).PokemonInstance(slt).status
statusturn(1) = player(index).PokemonInstance(slt).statusstun
turnsneed(1) = player(index).PokemonInstance(slt).turnsneed
'
patk(2) = TempPlayer(index).PokemonBattle.atk
pdef(2) = TempPlayer(index).PokemonBattle.def
spatk(2) = TempPlayer(index).PokemonBattle.spatk
spdef(2) = TempPlayer(index).PokemonBattle.spdef
spd(2) = TempPlayer(index).PokemonBattle.spd
hp(2) = TempPlayer(index).PokemonBattle.hp
MaxHp(2) = TempPlayer(index).PokemonBattle.MaxHp
status(2) = TempPlayer(index).PokemonBattle.status
statusturn(2) = TempPlayer(index).PokemonBattle.statusturn
turnsneed(2) = TempPlayer(index).PokemonBattle.turnsneed
'
aname = Trim$(Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Name)
defname = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)
pp = player(index).PokemonInstance(slt).moves(mslot).pp
If Trim$(PokemonMove(move).Category) = "Other Damage" Then

End If
If Trim$(PokemonMove(move).Category) = "Physical Damage" Then
atk = patk(1)
def = pdef(2)
End If
If Trim$(PokemonMove(move).Category) = "Special Damage" Then
atk = spatk(1)
def = spdef(2)
End If
'If pp < 1 Then Exit Sub Disabled cause of struggle , there is already enough check!

atype = Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Type
atype2 = Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Type2
deftype1 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type
deftype2 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type2
alvl = player(index).PokemonInstance(slt).level

Case False
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
patk(2) = player(index).PokemonInstance(slt).batk
pdef(2) = player(index).PokemonInstance(slt).bdef
spatk(2) = player(index).PokemonInstance(slt).bspatk
spdef(2) = player(index).PokemonInstance(slt).bspdef
spd(2) = player(index).PokemonInstance(slt).bspd
hp(2) = player(index).PokemonInstance(slt).hp
MaxHp(2) = player(index).PokemonInstance(slt).MaxHp
status(2) = player(index).PokemonInstance(slt).status
statusturn(2) = player(index).PokemonInstance(slt).statusstun
turnsneed(2) = player(index).PokemonInstance(slt).turnsneed
'

patk(1) = TempPlayer(index).PokemonBattle.atk
pdef(1) = TempPlayer(index).PokemonBattle.def
spatk(1) = TempPlayer(index).PokemonBattle.spatk
spdef(1) = TempPlayer(index).PokemonBattle.spdef
spd(1) = TempPlayer(index).PokemonBattle.spd
hp(1) = TempPlayer(index).PokemonBattle.hp
MaxHp(1) = TempPlayer(index).PokemonBattle.MaxHp
status(1) = TempPlayer(index).PokemonBattle.status
statusturn(1) = TempPlayer(index).PokemonBattle.statusturn
turnsneed(1) = TempPlayer(index).PokemonBattle.turnsneed
'

defname = Trim$(Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Name)
aname = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)

miss = doesMiss(PokemonMove(move).accuracy)

If Trim$(PokemonMove(move).Category) = "Other Damage" Then

End If
If Trim$(PokemonMove(move).Category) = "Physical Damage" Then
atk = patk(1)
def = pdef(2)
End If
If Trim$(PokemonMove(move).Category) = "Special Damage" Then
atk = spatk(1)
def = spdef(2)
End If
atype = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type
atype2 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type2
deftype1 = Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Type
deftype2 = Pokemon(player(index).PokemonInstance(slt).PokemonNumber).Type2
alvl = TempPlayer(index).PokemonBattle.level
End Select














'MOVES

If miss = True Then
SendBattleMessage index, aname & "->" & defname & " used " & Trim$(PokemonMove(move).Name) & "!", BrightRed
SendBattleMessage index, aname & "->" & defname & " missed!", Black
used = True
End If
'................................
If Trim$(PokemonMove(move).Name) = "Tackle" Then
If miss = False Then
n = DealDemage(index, move, mtype, atk, def, critical, atype, atype2, alvl, deftype1, deftype2)
hp(2) = hp(2) - n
SendBattleMessage index, aname & "->" & defname & " used Tackle!", BrightRed
SendBattleMessage index, aname & "->" & defname & " dealt " & n & " damage!", Black
used = True
End If
End If
'.................................
If move = 135 Then 'STRUGGLE
If miss = False Then
Dim fgt As Long
n = DealDemage(index, 135, "None", atk, def, critical, atype, atype2, alvl, deftype1, deftype2)
hp(2) = hp(2) - n
fgt = (MaxHp(1) / 4)
hp(1) = hp(1) - fgt
SendBattleMessage index, aname & "->" & defname & " used Struggle!", BrightRed
SendBattleMessage index, aname & "->" & defname & " dealt " & n & " damage!", Black
SendBattleMessage index, aname & " lost " & fgt & " HP!", Black
used = True
End If
End If
'.................................

'CHECK FOR ADDITIONAL EFFECTS
If used = False Then
Call MoveAdditionalEffect(index, move, slt, OnOpponent) 'ADDITIONAL EFFECTS FUNCTION CODES
used = True

'Basically now we have checked for additional effect and data has been written to player temp lets check it
If TempPlayer(index).moveUsageTemp.damageInflict > 0 Then
hp(2) = hp(2) - TempPlayer(index).moveUsageTemp.damageInflict
SendBattleMessage index, aname & "->" & defname & " dealt " & TempPlayer(index).moveUsageTemp.damageInflict & " damage!", Black
End If
critical = TempPlayer(index).moveUsageTemp.isCritical
If TempPlayer(index).moveUsageTemp.customAttackUsage <> "" Then
If GetMoveID(TempPlayer(index).moveUsageTemp.customAttackUsage) > 0 Then
UseAnotherMove index, OnOpponent, TempPlayer(index).moveUsageTemp.customAttackUsage, slt, mslot
SendBattleMessage index, aname & " is using a " & Trim$(TempPlayer(index).moveUsageTemp.customAttackUsage), Black
Exit Sub
End If
End If


If TempPlayer(index).moveUsageTemp.fleeBattle = True Then
If TempPlayer(index).inNPCBattle = False Then
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
Call SendBattleInfo(index, 0, 3, 0)
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slt, YES
End If
End If

If TempPlayer(index).moveUsageTemp.powerSet > 0 Then
n = DealDemage(index, move, mtype, atk, def, critical, atype, atype2, alvl, deftypE, deftype2, False, 1, True, TempPlayer(index).moveUsageTemp.powerSet, 0)
Else
If TempPlayer(index).moveUsageTemp.powerModifier > 0 Then
n = DealDemage(index, move, mtype, atk, def, critical, atype, atype2, alvl, deftypE, deftype2, False, 1, False, 0, TempPlayer(index).moveUsageTemp.powerModifier)
SendBattleMessage index, aname & "->" & defname & " ,power changed by " & TempPlayer(index).moveUsageTemp.powerModifier & "%", Black
Else
n = DealDemage(index, move, mtype, atk, def, critical, atype, atype2, alvl, deftype1, deftype2)
End If
End If
'AFTER DAMAGE
Dim gainedHP As Long
If TempPlayer(index).moveUsageTemp.HPDamageModifier > 0 Then
gainedHP = (n * (TempPlayer(index).moveUsageTemp.HPDamageModifier / 100))
hp(1) = hp(1) + (n * (TempPlayer(index).moveUsageTemp.HPDamageModifier / 100))
SendBattleMessage index, aname & " gained " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.HPModifier > 0 Then
gainedHP = (hp(1) * (TempPlayer(index).moveUsageTemp.HPModifier / 100))
hp(1) = hp(1) + (hp(1) * (TempPlayer(index).moveUsageTemp.HPModifier / 100))
SendBattleMessage index, aname & " gained " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.HPTotalModifier > 0 Then
gainedHP = (MaxHp(1) * (TempPlayer(index).moveUsageTemp.HPTotalModifier / 100))
hp(1) = hp(1) + (MaxHp(1) * (TempPlayer(index).moveUsageTemp.HPTotalModifier / 100))
SendBattleMessage index, aname & " gained " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.recoilHPCurrentModifier > 0 Then
gainedHP = (hp(1) * (TempPlayer(index).moveUsageTemp.recoilHPCurrentModifier / 100))
hp(1) = hp(1) - (hp(1) * (TempPlayer(index).moveUsageTemp.recoilHPCurrentModifier / 100))
SendBattleMessage index, aname & " lost " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.recoilHPDamageModifier > 0 Then
gainedHP = (n * (TempPlayer(index).moveUsageTemp.recoilHPDamageModifier / 100))
hp(1) = hp(1) - (n * (TempPlayer(index).moveUsageTemp.recoilHPDamageModifier / 100))
SendBattleMessage index, aname & " lost " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.recoilHPTotalModifier > 0 Then
gainedHP = (MaxHp(1) * (TempPlayer(index).moveUsageTemp.recoilHPTotalModifier / 100))
hp(1) = hp(1) - (MaxHp(1) * (TempPlayer(index).moveUsageTemp.recoilHPTotalModifier / 100))
SendBattleMessage index, aname & " lost " & gainedHP & " HP", Black
End If
If TempPlayer(index).moveUsageTemp.resetStats = YES Or TempPlayer(index).moveUsageTemp.resetStatsMe = YES Then
player(index).PokemonInstance(slt).batk = player(index).PokemonInstance(slt).atk
player(index).PokemonInstance(slt).bdef = player(index).PokemonInstance(slt).def
player(index).PokemonInstance(slt).bspatk = player(index).PokemonInstance(slt).spatk
player(index).PokemonInstance(slt).bspdef = player(index).PokemonInstance(slt).spdef
player(index).PokemonInstance(slt).bspd = player(index).PokemonInstance(slt).spd
SendBattleMessage index, "All stat changes have been reset!", Red
End If
If status(1) = TempPlayer(index).moveUsageTemp.attackerStatus Or TempPlayer(index).moveUsageTemp.attackerStatus < STATUS_NOTHING Then
Else
status(1) = TempPlayer(index).moveUsageTemp.attackerStatus
If TempPlayer(index).moveUsageTemp.attackerStatusRounds > 0 Then
statusturn(1) = 0
turnsneed(1) = TempPlayer(index).moveUsageTemp.attackerStatusRounds
Else
statusturn(1) = 0
If TempPlayer(index).moveUsageTemp.attackerStatus = STATUS_FLINCHED Then
turnsneed(1) = 1
Else
turnsneed(1) = Rand(2, 4)
End If
End If
End If
If status(2) = TempPlayer(index).moveUsageTemp.defenderStatus Or TempPlayer(index).moveUsageTemp.defenderStatus < STATUS_NOTHING Then
Else
status(2) = TempPlayer(index).moveUsageTemp.defenderStatus
If TempPlayer(index).moveUsageTemp.defenderStatusRounds > 0 Then
statusturn(2) = 0
turnsneed(2) = TempPlayer(index).moveUsageTemp.defenderStatusRounds
Else
statusturn(2) = 0
If TempPlayer(index).moveUsageTemp.defenderStatus = STATUS_FLINCHED Then
turnsneed(2) = 1
Else
turnsneed(2) = Rand(2, 4)
End If
End If
End If
'deal the actual damage
If Trim$(PokemonMove(move).Category) <> "Other Damage" Then
Dim xa As Double
Dim yA As Double
Dim typeModifier As Long
xa = GetTypeEffect(MoveTypes(move), deftype1)
yA = GetTypeEffect(MoveTypes(move), deftype2)
typeModifier = (xa * yA) * 100
If TempPlayer(index).moveUsageTemp.multiHit > 0 Then
For i = 1 To TempPlayer(index).moveUsageTemp.multiHit
hp(2) = hp(2) - n
Next
If n > 0 Then
SendBattleMessage index, aname & "->" & defname & " dealt " & n & " damage!", Black
SendBattleMessage index, aname & "->" & defname & " hitted " & TempPlayer(index).moveUsageTemp.multiHit & " times.", Black
SendBattleMessage index, aname & "->" & defname & " is " & typeModifier & "% effective!", Black
End If
Else
hp(2) = hp(2) - n
If n > 0 Then
SendBattleMessage index, aname & "->" & defname & " dealt " & n & " damage!", Black
SendBattleMessage index, aname & "->" & defname & " is " & typeModifier & "% effective!", Black
End If
End If
End If






















End If
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Select Case OnOpponent
Case True
If hp(1) < 0 Then
hp(1) = 0
End If
If hp(1) > player(index).PokemonInstance(slt).MaxHp Then
hp(1) = player(index).PokemonInstance(slt).MaxHp
End If
If hp(2) < 0 Then
hp(2) = 0
End If
If hp(2) > TempPlayer(index).PokemonBattle.MaxHp Then
hp(2) = TempPlayer(index).PokemonBattle.MaxHp
End If
player(index).PokemonInstance(slt).hp = hp(1)
player(index).PokemonInstance(slt).batk = patk(1)
player(index).PokemonInstance(slt).bdef = pdef(1)
player(index).PokemonInstance(slt).bspatk = spatk(1)
player(index).PokemonInstance(slt).bspdef = spdef(1)
player(index).PokemonInstance(slt).bspd = spd(1)
player(index).PokemonInstance(slt).status = status(1)
player(index).PokemonInstance(slt).statusstun = statusturn(1)
player(index).PokemonInstance(slt).turnsneed = turnsneed(1)
TempPlayer(index).PokemonBattle.hp = hp(2)
TempPlayer(index).PokemonBattle.atk = patk(2)
TempPlayer(index).PokemonBattle.def = pdef(2)
TempPlayer(index).PokemonBattle.spatk = spatk(2)
TempPlayer(index).PokemonBattle.spdef = spdef(2)
TempPlayer(index).PokemonBattle.spd = spd(2)
TempPlayer(index).PokemonBattle.status = status(2)
TempPlayer(index).PokemonBattle.statusturn = statusturn(2)
TempPlayer(index).PokemonBattle.turnsneed = turnsneed(2)
Case False
player(index).PokemonInstance(slt).hp = hp(2)
player(index).PokemonInstance(slt).batk = patk(2)
player(index).PokemonInstance(slt).bdef = pdef(2)
player(index).PokemonInstance(slt).bspatk = spatk(2)
player(index).PokemonInstance(slt).bspdef = spdef(2)
player(index).PokemonInstance(slt).bspd = spd(2)
player(index).PokemonInstance(slt).status = status(2)
player(index).PokemonInstance(slt).statusstun = statusturn(2)
player(index).PokemonInstance(slt).turnsneed = turnsneed(2)
TempPlayer(index).PokemonBattle.hp = hp(1)
TempPlayer(index).PokemonBattle.atk = patk(1)
TempPlayer(index).PokemonBattle.def = pdef(1)
TempPlayer(index).PokemonBattle.spatk = spatk(1)
TempPlayer(index).PokemonBattle.spdef = spdef(1)
TempPlayer(index).PokemonBattle.spd = spd(1)
TempPlayer(index).PokemonBattle.status = status(1)
TempPlayer(index).PokemonBattle.statusturn = statusturn(1)
TempPlayer(index).PokemonBattle.turnsneed = turnsneed(1)
End Select
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
SendBattleUpdate index, slt, YES

End Sub
































Sub CustomPoke(ByVal index As Long, ByVal pokeNum As Long, ByVal pokeLevel As Long, ByVal isPokeShiny As Long)
On Error Resume Next
If TempPlayer(index).PokemonBattle.PokemonNumber > 0 Then Exit Sub
ResetBattlePokemon (index)

Dim i As Long
Dim x As Long
Dim wildpoke As Long
Dim slot As Long
Dim frmlvl As Long
Dim tolvl As Long
Dim cstm As Long
Dim slt As Long
If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If

If slot < 1 Then Exit Sub


'Set wild pokemon

wildpoke = pokeNum



If wildpoke < 1 Or wildpoke > 721 Then Exit Sub 'No battle if there is not pokemon to spawn

'If there is pokemon then we are going to set BattlePokemon ready!

ResetBattlePokemon (index)
TempPlayer(index).PokemonBattle.PokemonNumber = wildpoke
TempPlayer(index).PokemonBattle.level = pokeLevel
TempPlayer(index).PokemonBattle.MapSlot = slt
TempPlayer(index).PokemonBattle.nature = Rand(1, MAX_NATURES)
TempPlayer(index).PokemonBattle.status = STATUS_NOTHING
TempPlayer(index).PokemonBattle.turnsneed = 0
TempPlayer(index).PokemonBattle.statusturn = 0
TempPlayer(index).PokemonBattle.isShiny = isPokeShiny
TempPlayer(index).PokemonBattle.atk = CalculateStat(Pokemon(wildpoke).atk, STAT_ATK)
TempPlayer(index).PokemonBattle.def = CalculateStat(Pokemon(wildpoke).def, STAT_DEF)
TempPlayer(index).PokemonBattle.spatk = CalculateStat(Pokemon(wildpoke).spatk, STAT_SPATK)
TempPlayer(index).PokemonBattle.spdef = CalculateStat(Pokemon(wildpoke).spdef, STAT_SPDEF)
TempPlayer(index).PokemonBattle.spd = CalculateStat(Pokemon(wildpoke).spd, STAT_SPEED)
TempPlayer(index).PokemonBattle.MaxHp = CalculateStat(Pokemon(wildpoke).MaxHp, STAT_HP)

If pokeLevel > 1 Then
Dim availableTP As Long
availableTP = pokeLevel * 3 - 3
Do While availableTP = 0
Dim stattoadd As Long
stattoadd = Rand(1, 6)
Select Case stattoadd
Case STAT_ATK
TempPlayer(index).PokemonBattle.atk = TempPlayer(index).PokemonBattle.atk + 1
Case STAT_DEF
TempPlayer(index).PokemonBattle.def = TempPlayer(index).PokemonBattle.def + 1
Case STAT_SPATK
TempPlayer(index).PokemonBattle.spatk = TempPlayer(index).PokemonBattle.spatk + 1
Case STAT_SPDEF
TempPlayer(index).PokemonBattle.spdef = TempPlayer(index).PokemonBattle.spdef + 1
Case STAT_SPEED
TempPlayer(index).PokemonBattle.spd = TempPlayer(index).PokemonBattle.spd + 1
Case STAT_HP
TempPlayer(index).PokemonBattle.MaxHp = TempPlayer(index).PokemonBattle.MaxHp + 2
End Select
availableTP = availableTP - 1
Loop
End If
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.MaxHp




For x = 1 To 6
player(index).PokemonInstance(x).batk = player(index).PokemonInstance(x).atk
player(index).PokemonInstance(x).bdef = player(index).PokemonInstance(x).def
player(index).PokemonInstance(x).bspd = player(index).PokemonInstance(x).spd
player(index).PokemonInstance(x).bspatk = player(index).PokemonInstance(x).spatk
player(index).PokemonInstance(x).bspdef = player(index).PokemonInstance(x).spdef
Next

'Set turn (My Speed>Enemy Speed = MyTurn)
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
TempPlayer(index).BattleCurrentTurn = 1
SendNpcBattle index, slot
'Call SendActionMsg(GetPlayerMap(index), "Encounter: " & Trim$(Pokemon(wildpoke).Name), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32)

End Sub


Function CheckItem(ByVal index As Long, ByVal iteminv As Long, ByVal slt As Long) As Boolean
If TempPlayer(index).inNPCBattle = True Then Exit Function
On Error Resume Next
Dim itemNum As Long
itemNum = GetPlayerInvItemNum(index, iteminv)
If itemNum < 1 Then Exit Function
If item(itemNum).Type = ITEM_TYPE_POKEBALL Then
CheckItem = True
If isCatched(index, TempPlayer(index).PokemonBattle.MaxHp, TempPlayer(index).PokemonBattle.hp, item(itemNum).CatchRate) Then
CatchPokemon index, slt
Else
End If
TakeItem index, itemNum, 1
End If
If itemNum < 1 Then Exit Function
If item(itemNum).Type = ITEM_TYPE_POKEPOTION Then
CheckItem = True
Call UsePotion(index, slt, item(itemNum).AddHP)
TakeItem index, itemNum, 1
End If

End Function

Sub CatchPokemon(ByVal index As Long, ByVal myslot As Long)
On Error Resume Next
Dim freeslot As Long
Dim freestorageslot As Long
Dim i As Long

For i = 1 To 6
If player(index).PokemonInstance(i).PokemonNumber = 0 Then
freeslot = i
Exit For
End If
Next

If Not freeslot > 0 Then
For i = 1 To 250
If player(index).StoragePokemonInstance(i).PokemonNumber = 0 Then
freestorageslot = i
Exit For
End If
Next
End If

Dim a As Long
Dim b As Long
Dim l As Long
For i = 1 To 4
If TempPlayer(index).PokemonBattle.moves(i).number = 0 Then
a = a + 1
End If
Next

If a = 0 Then
TempPlayer(index).PokemonBattle.moves(1).number = 1
End If

If freeslot > 0 Then
player(index).PokemonInstance(freeslot).PokemonNumber = TempPlayer(index).PokemonBattle.PokemonNumber
player(index).PokemonInstance(freeslot).MaxHp = TempPlayer(index).PokemonBattle.MaxHp
player(index).PokemonInstance(freeslot).hp = TempPlayer(index).PokemonBattle.hp
player(index).PokemonInstance(freeslot).level = TempPlayer(index).PokemonBattle.level
player(index).PokemonInstance(freeslot).nature = TempPlayer(index).PokemonBattle.nature
'MOVE ADDING

  
            For a = 1 To 4
            If GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a) > 0 Then
            player(index).PokemonInstance(freeslot).moves(a).number = GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a)
            player(index).PokemonInstance(freeslot).moves(a).pp = PokemonMove(GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a)).pp
            Else
            player(index).PokemonInstance(freeslot).moves(a).number = 0
            player(index).PokemonInstance(freeslot).moves(a).pp = 0
            End If
            Next
         
      
           
'MOVE ADDING
player(index).PokemonInstance(freeslot).atk = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).atk, STAT_ATK)
player(index).PokemonInstance(freeslot).def = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).def, STAT_DEF)
player(index).PokemonInstance(freeslot).spdef = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spdef, STAT_SPDEF)
player(index).PokemonInstance(freeslot).spatk = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spatk, STAT_SPATK)
player(index).PokemonInstance(freeslot).spd = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spd, STAT_SPEED)
player(index).PokemonInstance(freeslot).isShiny = TempPlayer(index).PokemonBattle.isShiny
player(index).PokemonInstance(freeslot).isTradeable = YES
If player(index).PokemonInstance(freeslot).isShiny = YES Then
player(index).PokemonInstance(freeslot).TP = 20
Else
player(index).PokemonInstance(freeslot).TP = 0
End If
player(index).PokemonInstance(freeslot).EXP = 0
'-----------------------------------------------------------------------------
If player(index).PokemonInstance(freeslot).level > 1 Then 'Add random stats
Dim availableTP As Long
availableTP = player(index).PokemonInstance(freeslot).level * 3 - 3
Do While availableTP > 0
Dim stattoadd As Long
stattoadd = Rand(2, 6)
Select Case stattoadd
Case STAT_ATK
player(index).PokemonInstance(freeslot).atk = player(index).PokemonInstance(freeslot).atk + 1
Case STAT_DEF
player(index).PokemonInstance(freeslot).def = player(index).PokemonInstance(freeslot).def + 1
Case STAT_SPATK
player(index).PokemonInstance(freeslot).spatk = player(index).PokemonInstance(freeslot).spatk + 1
Case STAT_SPDEF
player(index).PokemonInstance(freeslot).spdef = player(index).PokemonInstance(freeslot).spdef + 1
Case STAT_SPEED
player(index).PokemonInstance(freeslot).spd = player(index).PokemonInstance(freeslot).spd + 1
'Case STAT_HP
'player(index).PokemonInstance(freeslot).atk = player(index).PokemonInstance(freeslot).atk + 2
End Select
availableTP = availableTP - 1
Loop
End If
'-------------------------------------------------------------------------------

FinishWild index, myslot
PlayerMsg index, "Pokemon caught!", Yellow
Exit Sub
End If
If freestorageslot > 0 Then
player(index).StoragePokemonInstance(freestorageslot).PokemonNumber = TempPlayer(index).PokemonBattle.PokemonNumber
player(index).StoragePokemonInstance(freestorageslot).MaxHp = TempPlayer(index).PokemonBattle.MaxHp
player(index).StoragePokemonInstance(freestorageslot).hp = TempPlayer(index).PokemonBattle.hp
player(index).StoragePokemonInstance(freestorageslot).level = TempPlayer(index).PokemonBattle.level
player(index).StoragePokemonInstance(freestorageslot).nature = TempPlayer(index).PokemonBattle.nature
'MOVES

            For a = 1 To 4
            If GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a) > 0 Then
            player(index).StoragePokemonInstance(freestorageslot).moves(a).number = GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a)
            player(index).StoragePokemonInstance(freestorageslot).moves(a).pp = PokemonMove(GetPokemonMove(TempPlayer(index).PokemonBattle.PokemonNumber, a)).pp
            Else
            player(index).StoragePokemonInstance(freestorageslot).moves(a).number = 0
            player(index).StoragePokemonInstance(freestorageslot).moves(a).pp = 0
            End If
            Next
'MOVEs
player(index).StoragePokemonInstance(freestorageslot).atk = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).atk, STAT_ATK)
player(index).StoragePokemonInstance(freestorageslot).def = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).def, STAT_DEF)
player(index).StoragePokemonInstance(freestorageslot).spdef = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spdef, STAT_SPDEF)
player(index).StoragePokemonInstance(freestorageslot).spatk = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spatk, STAT_SPATK)
player(index).StoragePokemonInstance(freestorageslot).spd = CalculateStat(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).spd, STAT_SPEED)
player(index).StoragePokemonInstance(freestorageslot).isShiny = TempPlayer(index).PokemonBattle.isShiny
player(index).StoragePokemonInstance(freestorageslot).isTradeable = YES
If player(index).StoragePokemonInstance(freestorageslot).isShiny = YES Then
player(index).StoragePokemonInstance(freestorageslot).TP = 20
Else
player(index).StoragePokemonInstance(freestorageslot).TP = 0
End If

player(index).StoragePokemonInstance(freestorageslot).EXP = 0
'------------------------------------------------------------
If player(index).StoragePokemonInstance(freestorageslot).level > 1 Then 'Add random stats
Dim availableTP2 As Long
availableTP2 = player(index).StoragePokemonInstance(freestorageslot).level * 3 - 3
Do While availableTP2 > 0
Dim stattoadd2 As Long
stattoadd2 = Rand(2, 6)
Select Case stattoadd2
Case STAT_ATK
player(index).StoragePokemonInstance(freestorageslot).atk = player(index).StoragePokemonInstance(freestorageslot).atk + 1
Case STAT_DEF
player(index).StoragePokemonInstance(freestorageslot).def = player(index).StoragePokemonInstance(freestorageslot).def + 1
Case STAT_SPATK
player(index).StoragePokemonInstance(freestorageslot).spatk = player(index).StoragePokemonInstance(freestorageslot).spatk + 1
Case STAT_SPDEF
player(index).StoragePokemonInstance(freestorageslot).spdef = player(index).StoragePokemonInstance(freestorageslot).spdef + 1
Case STAT_SPEED
player(index).StoragePokemonInstance(freestorageslot).spd = player(index).StoragePokemonInstance(freestorageslot).spd + 1
'Case STAT_HP
'player(index).PokemonInstance(freeslot).atk = player(index).PokemonInstance(freeslot).atk + 2
End Select
availableTP2 = availableTP2 - 1
Loop
End If
'------------------------------------------------------------

FinishWild index, myslot
PlayerMsg index, "Pokemon stored in storage!", Yellow
PlayerMsg index, "Slot #" & freestorageslot, Yellow
Exit Sub
End If
End Sub

Function isCatched(ByVal index As Long, ByVal MaxHp As Long, ByVal hp As Long, ByVal CatchRate As Long) As Boolean
On Error Resume Next
If TempPlayer(index).inNPCBattle = True Then Exit Function
If TempPlayer(index).isCatchable = 1 Then Exit Function
Dim a As Long
Dim m As Long
'a = (MaxHp * 255 * 4) / (Hp * CatchRate) * (MaxHp * 2 / (Hp * 1.5))
a = ((3 * MaxHp - 2 * hp) * CatchRate) / (3 * MaxHp)
m = Rand(0, 255)
If player(index).Access >= 3 Then
PlayerMsg index, "[Admin Stat]Catch number:" & a & "," & m, Yellow
End If
If a >= m Then
isCatched = True
Else
isCatched = False
End If
End Function

Function Runned(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
If TempPlayer(index).inNPCBattle = True Then Exit Function
Dim F As Long
Dim a As Long
Dim b As Long
Dim C As Long

a = player(index).PokemonInstance(slot).spd
b = TempPlayer(index).PokemonBattle.spd 'Mod 256
If b > 0 Then
F = (((a * 128) / b) + 30) 'Mod 256
C = Rand(0, 255)
If GetPlayerAccess(index) >= ADMIN_DEVELOPER Then
PlayerMsg index, "[ADMIN STAT] RUN: " & F & "," & C, Yellow
End If
If F > C Then
Runned = True
Else
Runned = False
End If
Else
Runned = True
End If

End Function
Sub ResetBattlePokemon(ByVal index As Long)
On Error Resume Next
Dim i As Long
TempPlayer(index).PokemonBattle.PokemonNumber = 0
TempPlayer(index).PokemonBattle.atk = 0
TempPlayer(index).PokemonBattle.def = 0
TempPlayer(index).PokemonBattle.hp = 0
TempPlayer(index).PokemonBattle.isAttracted = 0
TempPlayer(index).PokemonBattle.isBurned = 0
TempPlayer(index).PokemonBattle.isConfused = 0
TempPlayer(index).PokemonBattle.isCursed = 0
TempPlayer(index).PokemonBattle.isFlinched = 0
TempPlayer(index).PokemonBattle.isFreezed = 0
TempPlayer(index).PokemonBattle.isParalized = 0
TempPlayer(index).PokemonBattle.isPoisoned = 0
TempPlayer(index).PokemonBattle.isShiny = 0
TempPlayer(index).PokemonBattle.isSleeping = 0
TempPlayer(index).PokemonBattle.level = 0
TempPlayer(index).PokemonBattle.MapSlot = 0
TempPlayer(index).PokemonBattle.MaxHp = 0
TempPlayer(index).PokemonBattle.nature = 0
TempPlayer(index).lastPoke = TempPlayer(index).PokemonBattle.PokemonNumber
TempPlayer(index).PokemonBattle.PokemonNumber = 0
TempPlayer(index).PokemonBattle.spatk = 0
TempPlayer(index).PokemonBattle.spd = 0
TempPlayer(index).PokemonBattle.spdef = 0
TempPlayer(index).PokemonBattle.status = 0
TempPlayer(index).PokemonBattle.statusturn = 0
TempPlayer(index).PokemonBattle.turnsneed = 0
TempPlayer(index).PokemonBattle.FirstMove = 0
For i = 1 To 4
TempPlayer(index).PokemonBattle.moves(i).number = 0
Next
End Sub


Sub FinishWild(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
exp_gained = GetExpWildBattle(index, slot)
Dim pc As Long
pc = Rand(1, 3)
If isPlayerMember(index) Then
pc = pc * 1.5
End If
'GiveItem index, 1, pc
Call GiveEXPtoPlayer(index, slot, exp_gained)
'Close the battle
Call GiveItem(index, 1, pc, NO)
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
Call SendBattleInfo(index, pc, 4, exp_gained)
ResetBattlePokemon (index)
SendIsInBattle index
SendPlayerPokemon index
SendBattleUpdate index, slot
'SendSound index, "Victory.wav"
End Sub
Sub UsePotion(ByVal index As Long, ByVal slot As Long, ByVal heal As Long)
On Error Resume Next
If player(index).PokemonInstance(slot).PokemonNumber < 1 Or player(index).PokemonInstance(slot).hp < 1 Then
Exit Sub
End If
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).hp + heal
If player(index).PokemonInstance(slot).hp > player(index).PokemonInstance(slot).MaxHp Then
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).MaxHp
End If
'done
End Sub
Function doesMiss(ByVal accuracy As Long) As Boolean
On Error Resume Next
Dim i As Long
i = Rand(1, 100)
If i > accuracy Then
doesMiss = True
Else
doesMiss = False
End If
End Function





Public Sub LoadMoveData(ByVal index As String, ByVal move As Long, ByVal slot As Long)
On Error Resume Next
Dim moveused As String
Dim pokeSlot As Long
moveused = Trim$(TempPlayer(index).moveData.move)
pokeNum = TempPlayer(index).moveData.PokemonNumber
ClearMoveData (index)
If moveused <> "" Then
If Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) = Trim$(Pokemon(pokeNum).Name) Then
TempPlayer(index).moveData.lastMove = moveused
End If
TempPlayer(index).moveData.move = Trim$(PokemonMove(move).Name)
TempPlayer(index).moveData.PokemonNumber = player(index).PokemonInstance(slot).PokemonNumber
End If
End Sub
Public Sub MoveAdditionalEffect(ByVal index As Long, ByVal move As Long, ByVal slot As Long, ByVal amI As Boolean)
On Error Resume Next
Dim usingMove As String
Dim message As String
Dim defenderStatus As Long
Dim attackerStatus As Long
Dim defenderstatusnew As Long
Dim attackerStatusNew As Long
Dim defenderStatusNewTurns As Long
Dim attackerStatusNewTurns As Long
Dim defenderName As String
Dim attackerName As String
Dim powerModifier As Long 'PERCENT
Dim defenderStats As StatsRec
Dim attackerstats As StatsRec
Dim msgColor As Long
Dim dmgAlready As Boolean
Dim reserstatsMe As Boolean
Dim resetStatsMeTo As Long
Dim resetStats As Boolean
Dim resetStatsTo As Long
Dim onlyOneMove As Boolean
Dim onlyOneMoveRounds As Long
Dim onlyOneMoveAdditionalEffect As String
Dim attackerTypeChange As Long
Dim defenderTypeChange As Long
Dim damageInflict As Long 'Nanosenje damagea
Dim attackersHP As Long
Dim defendersHP As Long
Dim attackersLevel As Long
Dim defendersLevel As Long
Dim attackersMaxHp As Long
Dim defendersMaxHp As Long
Dim attackersSpd As Long
Dim defendersSpd As Long
Dim hitTimes As Long 'MULTI HIT
'EFFECT
Dim effectBegin As String
Dim effectLast As Long
'
Dim HPModifier As Long 'This gives hp by current health
Dim HPDamageModifier As Long 'This gives hp by dealt damage
Dim HPTotalModifier As Long 'This gives hp by total hp
Dim powerSet As Long 'Sets power
Dim attackersPP As Long
Dim isCritical As Long
Call LoadMoveData(index, move, slot)
Dim lastAttack As String
Dim customAttackUsage As String

Dim recoilHPTotalModifier As Long
Dim recoilHPDamageModifier As Long
Dim recoilHPCurrentModifier As Long

Dim fleeBattle As Boolean
Dim attackerType1 As Long
Dim attackerType2 As Long
Dim defenderType1 As Long
Dim defenderType2 As Long

msgColor = Black 'Load default color
lastAttack = TempPlayer(index).moveData.lastMove
'load things


Dim fnCode As String
fnCode = getMoveFunctionCode(move)



usingMove = Trim$(PokemonMove(move).Name)
If amI = True Then
defenderStatus = TempPlayer(index).PokemonBattle.status
attackerStatus = player(index).PokemonInstance(slot).status
defenderName = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)
attackerName = Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name)
attackersHP = player(index).PokemonInstance(slot).hp
defendersHP = TempPlayer(index).PokemonBattle.hp
attackersLevel = player(index).PokemonInstance(slot).level
defendersLevel = TempPlayer(index).PokemonBattle.level
attackersMaxHp = player(index).PokemonInstance(slot).MaxHp
defendersMaxHp = TempPlayer(index).PokemonBattle.MaxHp
attackersSpd = player(index).PokemonInstance(slot).spd
defendersSpd = TempPlayer(index).PokemonBattle.spd
attackerType1 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type
attackerType2 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type2
defenderType1 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type
defenderType2 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type2
Else
defenderStatus = player(index).PokemonInstance(slot).status
attackerStatus = TempPlayer(index).PokemonBattle.status
defenderName = Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name)
attackerName = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)
attackersHP = TempPlayer(index).PokemonBattle.hp
defendersHP = player(index).PokemonInstance(slot).hp
attackersLevel = TempPlayer(index).PokemonBattle.level
defendersLevel = player(index).PokemonInstance(slot).level
attackersMaxHp = TempPlayer(index).PokemonBattle.MaxHp
defendersMaxHp = player(index).PokemonInstance(slot).MaxHp
defendersSpd = player(index).PokemonInstance(slot).spd
attackersSpd = TempPlayer(index).PokemonBattle.spd
defenderType1 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type
defenderType2 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type2
attackerType1 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type
attackerType2 = Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Type2
End If


Dim aX As Double
Dim aY As Double

Dim bXY As Double
aX = GetTypeEffect(MoveTypes(move), defenderType1)
aY = GetTypeEffect(MoveTypes(move), defenderType2)
If defenderType1 > 0 And defenderType2 > 0 Then
bXY = aX * aY
Else
If defenderType1 > 0 Then
bXY = aX
Else
If defenderType2 > 0 Then
bXY = aY
End If
End If
End If


Dim x As Long
Dim y As Long
'Start
Select Case fnCode ' LONG RIDE
Case "000" ' No effect
attack = YES
Case "001" 'Nothing happens at all
message = "Nothing happens at all"
Case "002"
'Struggle
'Already done
Case "003"
defenderstatusnew = STATUS_SLEEPING

Case "004"
defenderstatusnew = STATUS_SLEEPING

Case "005"
x = Rand(1, 100)
If usingMove = "Gunk Shot" Or usingMove = "Poison Gas" Or usingMove = "Poison Jab" Or usingMove = "Poison Sting" Or usingMove = "Poison Tail" Then
If x <= 10 Then
defenderstatusnew = STATUS_POISONED
End If
Else
If x <= 30 Then
defenderstatusnew = STATUS_POISONED
End If
End If
msgColor = Red
Case "006"
defenderstatusnew = STATUS_BADLYPOISONED

msgColor = Red
Case "007"

x = Rand(1, 100)
If usingMove = "Thunder Punch" Or usingMove = "Lick" Or usingMove = "Spark" Or usingMove = "Thunderbolt" Or usingMove = "Thunder Shock" Then
If x <= 10 Then
defenderstatusnew = STATUS_PARALIZED
End If
Else
If usingMove = "Body Slam" Or usingMove = "Discharge" Then
If x <= 30 Then
defenderstatusnew = STATUS_PARALIZED
End If
Else
defenderstatusnew = STATUS_PARALIZED
End If
End If

msgColor = Red
Case "008"
defenderstatusnew = STATUS_PARALIZED

msgColor = Red
Case "009"

x = Rand(1, 100)
y = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_PARALIZED
End If
If y <= 10 Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "00A"
If usingMove = "Blue Flare" Then
powerModifier = 200 '2 times
End If

x = Rand(1, 100)
If usingMove = "Inferno" Then
defenderstatusnew = STATUS_BURNED
Else
If usingMove = "Scald" Then
If x <= 30 Then
defenderstatusnew = STATUS_BURNED
End If
Else
If x <= 10 Then
defenderstatusnew = STATUS_BURNED
End If
End If
End If

Case "00B"
x = Rand(1, 100)
y = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_BURNED
End If
If y <= 10 Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "00C"

x = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_FREEZED
End If



Case "00D"

x = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_FREEZED
End If
Case "00E"
x = Rand(1, 100)
y = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_FREEZED
End If
If y <= 10 Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "00F"
x = Rand(1, 100)
If usingMove = "Hyper Fang" Or usingMove = "Extrasensory" Then
If x <= 10 Then
defenderstatusnew = STATUS_FLINCHED
End If
Else
If usingMove = "Dark Pulse" Or usingMove = "Zen Headbutt" Then
If x <= 20 Then
defenderstatusnew = STATUS_FLINCHED
End If
Else
If x <= 30 Then
defenderstatusnew = STATUS_FLINCHED
End If
End If
End If
Case "010"

x = Rand(1, 100)
If usingMove = "Stomp" Then
If x <= 30 Then
defenderstatusnew = STATUS_FLINCHED
End If
Else
defenderstatusnew = STATUS_FLINCHED
End If
Case "011"
If defenderStatus = STATUS_SLEEPING Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "012"
x = Rand(1, 100)
If x <= 30 Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "013"
defenderstatusnew = STATUS_CONFUSED

Case "014"
x = Rand(1, 100)
If x <= 25 Then
defenderstatusnew = STATUS_CONFUSED
End If

Case "015"
defenderstatusnew = STATUS_CONFUSED

Case "016"
'NOTHING

Case "017"

y = Rand(1, 100)
x = Rand(1, 3)
If y <= 10 Then
If x = 1 Then
defenderstatusnew = STATUS_BURNED
End If
If x = 2 Then
defenderstatusnew = STATUS_PARALIZED
End If
If x = 3 Then
defenderstatusnew = STATUS_FREEZED
End If
End If
Case "018"
If attackerStatus = STATUS_BURNED Or attackerStatus = STATUS_PARALIZED Or attackerStatus = STATUS_POISONED Then
attackerStatusNew = STATUS_NOTHING
End If

Case "019"
'NOTHING FOR NOW

Case "01A"
'NOTHING FOR NOW

Case "01B"
'NOTHING FOR NOW

Case "01C"
attackerstats.atk = attackerstats.atk + 1

Case "01D"
attackerstats.def = attackerstats.def + 1

Case "01E"
attackerstats.atk = attackerstats.atk + 1

Case "01F"
attackerstats.spd = attackerstats.spd + 1

Case "020"
attackerstats.spatk = attackerstats.spatk + 1

Case "021"
attackerstats.spdef = attackerstats.spdef + 1

Case "022"
'NOTHING FOR NOW
Case "023"
attackerstats.criticalHit = attackerstats.criticalHit + 1
Case "024"
attackerstats.atk = attackerstats.atk + 1
attackerstats.def = attackerstats.def + 1

Case "025"
attackerstats.atk = attackerstats.atk + 1
attackerstats.def = attackerstats.def + 1
attackerstats.accuracy = attackerstats.accuracy + 1

Case "026"
attackerstats.atk = attackerstats.atk + 1
attackerstats.spd = attackerstats.spd + 1

Case "027"
attackerstats.atk = attackerstats.atk + 1
attackerstats.spatk = attackerstats.spatk + 1

Case "028"
attackerstats.atk = attackerstats.atk + 1
attackerstats.spatk = attackerstats.spatk + 1

Case "029"
attackerstats.atk = attackerstats.atk + 1
attackerstats.accuracy = attackerstats.atk + 1

Case "02A"
attackerstats.def = attackerstats.def + 1
attackerstats.spdef = attackerstats.spdef + 1

Case "02B"
attackerstats.spatk = attackerstats.spatk + 1
attackerstats.spdef = attackerstats.spdef + 1
attackerstats.spd = attackerstats.spd + 1

Case "02C"
attackerstats.spatk = attackerstats.spatk + 1
attackerstats.spdef = attackerstats.spdef + 1

Case "02D"
attackerstats.atk = attackerstats.atk + 1
attackerstats.def = attackerstats.def + 1
attackerstats.spd = attackerstats.spd + 1
attackerstats.spatk = attackerstats.spatk + 1
attackerstats.spdef = attackerstats.spdef + 1

Case "02E"
attackerstats.atk = attackerstats.atk + 2

Case "02F"
attackerstats.def = attackerstats.def + 2

Case "030"
attackerstats.spd = attackerstats.spd + 2

Case "031"
attackerstats.spd = attackerstats.spd + 2

Case "032"
attackerstats.spatk = attackerstats.spatk + 2

Case "033"
attackerstats.spdef = attackerstats.spdef + 2

Case "034"
'NOTHING
Case "035"
attackerstats.def = attackerstats.def - 1
attackerstats.spdef = attackerstats.spdef - 1
attackerstats.atk = attackerstats.atk + 2
attackerstats.spatk = attackerstats.spatk + 2
attackerstats.spd = attackerstats.spd + 2

Case "036"
attackerstats.atk = attackerstats.atk + 1
attackerstats.spd = attackerstats.spd + 2

Case "037"
attackerstats.spd = attackerstats.spd + 2

Case "038"
attackerstats.def = attackerstats.def + 3

Case "039"
attackerstats.spatk = attackerstats.spatk + 3

Case "03A"
'Nothing
Case "03B"
attackerstats.atk = attackerstats.atk - 1
attackerstats.def = attackerstats.def - 1

Case "03C"
attackerstats.spdef = attackerstats.spdef - 1
attackerstats.def = attackerstats.def - 1

Case "03D"
attackerstats.spdef = attackerstats.spdef - 1
attackerstats.def = attackerstats.def - 1
attackerstats.spd = attackerstats.spd - 1

Case "03E"
attackerstats.spd = attackerstats.spd - 1

Case "03F"
attackerstats.spatk = attackerstats.spatk - 2

Case "040"
x = Rand(1, 100)
y = Rand(1, 100)
If x <= 50 Then
defenderStats.spatk = defenderStats.spatk + 1
End If
If y <= 50 Then
defenderstatusnew = STATUS_CONFUSED
End If

Case "041"
x = Rand(1, 100)
y = Rand(1, 100)
If x <= 50 Then
defenderStats.atk = defenderStats.atk + 2
End If
If y <= 50 Then
defenderstatusnew = STATUS_CONFUSED
End If

Case "042"
defenderStats.atk = defenderStats.atk - 1

Case "043"
defenderStats.def = defenderStats.def - 1

Case "044"
defenderStats.spd = defenderStats.spd - 1

Case "045"
defenderStats.spatk = defenderStats.spatk - 1

Case "046"
defenderStats.spdef = defenderStats.spdef - 1

Case "047"
defenderStats.accuracy = defenderStats.accuracy - 1

Case "048"
'fuck this shit nothing

Case "049"
'SAME SHIT AS 048

Case "04A"
defenderStats.atk = defenderStats.atk - 1
defenderStats.def = defenderStats.def - 1

Case "04B"
defenderStats.atk = defenderStats.atk - 2

Case "04C"
defenderStats.def = defenderStats.def - 2

Case "04D"
defenderStats.spd = defenderStats.spd - 2

Case "04E"
defenderStats.spatk = defenderStats.spatk - 2

Case "04F"
defenderStats.spdef = defenderStats.spdef - 2

Case "050"
resetStatsMe = True
resetStatsMeTo = 0

Case "051"
resetStats = True
resetStatsTo = 0

Case "052"
Dim atkAtk As Long
Dim atkspAtk As Long
atkAtk = attackerstats.atk
atkspAtk = attackerstats.spatk
attackerstats.atk = defenderStats.atk
attackerstats.spatk = defenderStats.spatk
defenderStats.atk = atkAtk
defenderStats.spatk = atkspdatk
message = "Battlers attack and sp. attack changes have been switched!"

Case "053"
Dim atkdef As Long
Dim atkspdef As Long
atkdef = attackerstats.def
atkspdef = attackerstats.spdef
attackerstats.def = defenderStats.def
attackerstats.spdef = defenderStats.spdef
defenderStats.def = atkdef
defenderStats.spdef = atkspdef
message = "Battlers defense and sp. Def changes have been switched!"

Case "054"
Dim attackerTempStats As StatsRec
attackerTempStats = attackerstats
attackerstats = defenderStats
defenderStats = attackerTempStats
message = "Battlers stats changes have been switched!"

Case "055"
attackerstats = defenderStats
message = attackerName & " copied " & defenderName & " stat changes!"

Case "056"
'NOTHING FOR NOW!

Case "057"
'NOTHING FOR NOW

Case "058"
'Nothing for now

Case "059"
'Nothing for now

Case "05A"
'Nothing for now

Case "05B"
'Nothing for now

Case "05C"
'Nothing for now

Case "05D"
'Nothing for now

Case "05E"
'Nothing for now

Case "05F"
'Nothing for now

Case "060"
'Nothing for now

Case "061"
'Nothing for now

Case "062"
'Nothing for now

Case "063"
'Nothing for now

Case "064"
'Nothing for now

Case "065"
'Nothing for now

Case "066"
'Nothing for now

Case "067"
'Nothing for now

Case "068"
'Nothing for now

Case "069"
'Nothing for now
Case "06A"
damageInflict = 20

Case "06B"
damageInflict = 20

Case "06C"
Dim halfOfHP As Long
halfOfHP = defendersHP / 2
damageInflict = halfOfHP

Case "06D"
damageInflict = attackersLevel

Case "06E"
Dim hpDif As Long
hpDif = defendersHP - attackersHP
If defendersHP >= hpDif Then
Else
damageInflict = hpDif
End If

Case "06F"
n = Rand(1, 100)
Dim iDmg As Long
iDmg = (attackersLevel) * (n + 50) / 100
damageInflict = iDmg

Case "070"
If attackersLevel > defendersLevel Then
damageInflict = defendersHP
End If

Case "071"
'Nothing for now

Case "072"
'Same

Case "073"
'same

Case "074"
'NO ALLYIES YET

Case "075"
'Nothing
Case "076"
Case "077"

Case "078"
defenderstatusnew = STATUS_FLINCHED

Case "079"
Case "07A"

Case "07B"
If defenderStatus = STATUS_POISONED Then
powerModifier = 200
End If

Case "07C"
If defenderStatus = STATUS_PARALIZED Then
powerModifier = 200
defenderstatusnew = STATUS_NOTHING
End If

Case "07D"
If defenderStatus = STATUS_SLEEPING Then
powerModifier = 200
defenderstatusnew = STATUS_NOTHING
End If

Case "07E"
If defenderStatus = STATUS_POISONED Or defenderStatus = STATUS_BURNED Or defenderStatus = STATUS_PARALIZED Then
powerModifier = 200
End If

Case "07F"
If defenderStatus <> STATUS_NOTHING Then
powerModifier = 200
End If

Case "080"
If defendersHP <= (defendersMaxHp / 2) Then
powerModifier = 200
End If

Case "081"
'Nothing for now

Case "082"
'Nothing for now

Case "083"
'Nothing for now
Case "084"
If attackersSpd > defendersSpd Then
Else
powerModifier = 200
End If

Case "085"
'NFN

Case "086"
powerModifier = 200

Case "087"
'NFN

Case "088"
'NFN

Case "089"
x = Rand(1, 100)
If x < 20 Then
powerModifier = Rand(50, 70)
Else
powerModifier = Rand(70, 150)
End If

Case "08A"
x = Rand(1, 100)
If x < 20 Then
powerModifier = Rand(50, 70)
Else
powerModifier = Rand(70, 150)
End If

Case "08B"
x = (attackersHP / attackersMaxHp) * 150
powerModifier = x

Case "08C"
x = (defendersHP / defendersMaxHp) * 120
powerModifier = x

Case "08D"
x = (defendersSpd / attackersSpd) * 25
powerModifier = x

Case "08E"
'NFN

Case "08F"
'NFN

Case "090"
'NFN

Case "091"
'NFN

Case "092"
powerModifier = Rand(100, 500)

Case "093"
'NFN

Case "094"
x = Rand(1, 100)
If x <= 20 Then
HPModifier = 25
End If
If x > 20 And x <= 60 Then
powerSet = 40
End If
If x > 60 And x <= 90 Then
powerSet = 80
End If
If x > 90 And x <= 100 Then
powerSet = 120
End If

Case "095"
x = Rand(1, 100)
If x <= 5 Then
powerSet = 10
End If
If x > 5 And x <= 15 Then
powerSet = 30
End If
If x > 15 And x <= 35 Then
powerSet = 50
End If
If x > 35 And x <= 65 Then
powerSet = 70
End If
If x > 65 And x <= 85 Then
powerSet = 90
End If
If x > 85 And x <= 95 Then
powerSet = 110
End If
If x > 95 And x <= 100 Then
powerSet = 150
End If

Case "096"
Case "097"
powerSet = Rand(40, 200)

Case "098"
x = 48 * (attackersHP / attackersMaxHp)
If x >= 0 And x <= 1 Then
powerSet = 200
End If
If x >= 2 And x <= 4 Then
powerSet = 150
End If
If x >= 5 And x <= 9 Then
powerSet = 100
End If
If x >= 10 And x <= 16 Then
powerSet = 80
End If
If x >= 17 And x <= 32 Then
powerSet = 40
End If
If x >= 33 Then
powerSet = 20
End If

Case "099"
x = (attackersSpd / defendersSpd)
If x >= 4 Then
powerSet = 150
End If
If x = 3 Then
powerSet = 120
End If
If x = 2 Then
powerSet = 80
End If
If x = 1 Then
powerSet = 60
End If
If x < 1 Then
powerSet = 40
End If

Case "09A"
powerSet = Rand(20, 120)

Case "09B"
powerSet = Rand(40, 120)

Case "09C"
'NFN

Case "09D"
effectBegin = "MUDSPORT"
effectLast = 5

Case "09E"
effectBegin = "WATERSPORT"
effectLast = 5
Case "09F"
'NFN

Case "0A0"
isCritical = YES

Case "0A1"
'NFN

Case "0A2"
'NFN
Case "0A3"
'NFN
Case "0A4"
defenderstatusnew = STATUS_SLEEPING

Case "0A5"
'no need
Case "0A6"
'NFN
Case "0A7"
'NFN
Case "0A8"
'NFN
Case "0A9"
'NFN
Case "0AA"
'NFN
Case "0AB"
'NFN
Case "0AC"
'NFN
Case "0AD"
'NFN
Case "0AE"
If amI Then
customAttackUsage = lastAttack
End If

Case "0AF"
customAttackUsage = lastAttack

Case "0B0"
'NFN

Case "0B1"
'NFN
Case "0B3"
'nfn
Case "0B4"
Dim foundMoveSleepTalk As Boolean
Dim moveRnd As Long
'Sleep talk
If attackerStatus = STATUS_SLEEPING Then
If amI Then
'
Do While foundMoveSleepTalk = False
moveRnd = Rand(1, 4)
If Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(moveRnd).number).Name) = Trim$(usingMove) Then
Else
foundMoveSleepTalk = True
customAttackUsage = Trim$(PokemonMove(player(index).PokemonInstance(slot).moves(moveRnd).number).Name)
End If
Loop
'
Else
Do While foundMoveSleepTalk = False
moveRnd = Rand(1, 4)
If Trim$(PokemonMove(TempPlayer(index).PokemonBattle.moves(moveRnd).number).Name) = Trim$(usingMove) Then
Else
foundMoveSleepTalk = True
customAttackUsage = Trim$(PokemonMove(TempPlayer(index).PokemonBattle.moves(moveRnd).number).Name)
End If
Loop
End If
End If

Case "0B5"
'NFN

Case "0B6"
x = Rand(1, MAX_MOVES)
customAttackUsage = Trim$(PokemonMove(x).Name)

Case "0B7"
'NFN
Case "0B8"
'NFN
Case "0B9"
'NFN
Case "0BA"
'NFN
Case "0BB"
'NFN
Case "0BC"
'NFN
Case "0BD"
hitTimes = 2
Case "0BE"
hitTimes = 2
x = Rand(1, 100)
If x <= 30 Then
defenderstatusnew = STATUS_POISONED
End If

Case "0BF"
hitTimes = 3

Case "0C0"
x = Rand(1, 100)
If x <= 33 Then
hitTimes = 2
End If
If x > 33 And x <= 66 Then
hitTimes = 3
End If
If x > 66 And x <= 82 Then
hitTimes = 4
End If
If x > 82 Then
hitTimes = 5
End If

Case "0C1"
x = Rand(1, 3)
hitTimes = x

Case "0C2"
'NO NEED

Case "0C3"
'NO NEED

Case "0C4"
'NO NEED

Case "0C5"
x = Rand(1, 100)
If x <= 33 Then
defenderstatusnew = STATUS_PARALIZED
End If

Case "0C6"
x = Rand(1, 100)
If x <= 33 Then
defenderstatusnew = STATUS_BURNED
End If

Case "0C7"
x = Rand(1, 100)
If x <= 33 Then
defenderstatusnew = STATUS_FLINCHED
End If

Case "0C8"
attackerstats.def = attackerstats.def + 1

Case "0C9"
'No Need

Case "0CA"
'No need

Case "0CB"
'NO need

Case "0CC"
x = Rand(1, 100)
If x <= 33 Then
defenderstatusnew = STATUS_PARALIZED
End If

Case "0CD"
'No need

Case "0CE"
'No need

Case "0CF"
effectBegin = "0CF"
x = Rand(5, 6)
effectLast = x

Case "0D0"
effectBegin = "0D0"
x = Rand(5, 6)
effectLast = x

Case "0D1"
'NFN

Case "0D2"
onlyOneMove = True
onlyOneMoveRounds = 3
onlyOneMoveAdditionalEffect = "PETALDANCE"


Case "0D3"
'NFN
Case "0D4"
'NFN

Case "0D5"
HPTotalModifier = 50

Case "0D6"
HPTotalModifier = 50

Case "0D7"
'NFN

Case "0D8"
HPTotalModifier = 50

Case "0D9"
HPTotalModifier = 100
attackerStatusNew = STATUS_SLEEPING
attackerStatusNewTurns = 3

Case "0DA"
HPTotalModifier = 7

Case "0DB"
'NFN

Case "0DC"
'NFN
Case "0DD"
HPDamageModifier = 50

Case "0DE"
If defenderStatus = STATUS_SLEEPING Then
HPDamageModifier = 50
End If

Case "0DF"
HPTotalModifier = 50

Case "0E0"
recoilHPTotalModifier = 100

Case "0E1"
recoilHPTotalModifier = 100
damageInflict = attackersHP

Case "0E2"
recoilHPTotalModifier = 100
defenderStats.atk = defenderStats.atk - 2
defenderStats.spatk = defenderStats.spatk - 2



Case "0E3"
'NFN
Case "0E4"
'NFN
Case "0E5"
'NFN
Case "0E6"
'NFN
Case "0E7"
'NFN
Case "0E8"
Case "0E9"



Case "0EA"
fleeBattle = True

Case "0EB"
If attackersLevel >= defendersLevel Then
fleeBattle = True
End If


Case "0EC"
'NFN
Case "0ED"
'NFN
Case "0EE"
'NFN
Case "0EF"
'NFN
Case "0F0"
'NFN
Case "0F9"
'NFN

Case "0FA"
recoilHPDamageModifier = 25

Case "0FB"

x = Rand(1, 100)
If usingMove = "Brave Bird" Then
If x <= 30 Then
defenderstatusnew = STATUS_FLINCHED
End If
End If
recoilHPDamageModifier = 33

Case "0FC"
recoilHPDamageModifier = 50

Case "0FD"
recoilHPDamageModifier = 33
x = Rand(1, 100)
If x < 33 Then
defenderstatusnew = STATUS_PARALIZED
End If

Case "0FE"
recoilHPDamageModifier = 33
x = Rand(1, 100)
If x < 33 Then
defenderstatusnew = STATUS_PARALIZED
End If

Case "0FF"
'NFN
Case "100"
'NFN
Case "101"
'NFN
Case "102"
'NFN
Case "103"
'NFN
Case "104"
'NFN
Case "105"
'NFN
Case "106"
'NFN
Case "107"
'NFN
Case "108"
'NFN
Case "109"
'NFN
Case "10A"
'NFN


Case "10B"
'NO NEED


Case "10C"
'NFN
Case "10D"
'NFN
Case "10E"
'NFN
Case "10F"
'NFN
Case "110"
'NFN
Case "111"
'NFN
Case "112"
'NFN

Case "134"
message = "Congratulations!"
msgColor = Red

Case "135"

x = Rand(1, 100)
If x <= 10 Then
defenderstatusnew = STATUS_FREEZED
End If
'MOVES FROM 113 to END DO NOT WORK except of 134

'DONE
End Select







If amI = True Then
player(index).PokemonInstance(slot).batk = player(index).PokemonInstance(slot).batk + (player(index).PokemonInstance(slot).batk * (attackerstats.atk * 0.15))
player(index).PokemonInstance(slot).bdef = player(index).PokemonInstance(slot).bdef + (player(index).PokemonInstance(slot).bdef * (attackerstats.def * 0.15))
player(index).PokemonInstance(slot).bspd = player(index).PokemonInstance(slot).bspd + (player(index).PokemonInstance(slot).bspd * (attackerstats.spd * 0.15))
player(index).PokemonInstance(slot).bspatk = player(index).PokemonInstance(slot).bspatk + (player(index).PokemonInstance(slot).bspatk * (attackerstats.spatk * 0.15))
player(index).PokemonInstance(slot).bspdef = player(index).PokemonInstance(slot).bspdef + (player(index).PokemonInstance(slot).bspdef * (attackerstats.spdef * 0.15))
TempPlayer(index).PokemonBattle.atk = TempPlayer(index).PokemonBattle.atk + (TempPlayer(index).PokemonBattle.atk * (defenderStats.atk * 0.15))
TempPlayer(index).PokemonBattle.def = TempPlayer(index).PokemonBattle.def + (TempPlayer(index).PokemonBattle.def * (defenderStats.def * 0.15))
TempPlayer(index).PokemonBattle.spd = TempPlayer(index).PokemonBattle.spd + (TempPlayer(index).PokemonBattle.spd * (defenderStats.spd * 0.15))
TempPlayer(index).PokemonBattle.spatk = TempPlayer(index).PokemonBattle.spatk + (TempPlayer(index).PokemonBattle.spatk * (defenderStats.spatk * 0.15))
TempPlayer(index).PokemonBattle.spdef = TempPlayer(index).PokemonBattle.spdef + (TempPlayer(index).PokemonBattle.spdef * (defenderStats.spdef * 0.15))
Else
player(index).PokemonInstance(slot).batk = player(index).PokemonInstance(slot).batk + (player(index).PokemonInstance(slot).batk * (defenderStats.atk * 0.15))
player(index).PokemonInstance(slot).bdef = player(index).PokemonInstance(slot).bdef + (player(index).PokemonInstance(slot).bdef * (defenderStats.def * 0.15))
player(index).PokemonInstance(slot).bspd = player(index).PokemonInstance(slot).bspd + (player(index).PokemonInstance(slot).bspd * (defenderStats.spd * 0.15))
player(index).PokemonInstance(slot).bspatk = player(index).PokemonInstance(slot).bspatk + (player(index).PokemonInstance(slot).bspatk * (defenderStats.spatk * 0.15))
player(index).PokemonInstance(slot).bspdef = player(index).PokemonInstance(slot).bspdef + (player(index).PokemonInstance(slot).bspdef * (defenderStats.spdef * 0.15))
TempPlayer(index).PokemonBattle.atk = TempPlayer(index).PokemonBattle.atk + (TempPlayer(index).PokemonBattle.atk * (attackerstats.atk * 0.15))
TempPlayer(index).PokemonBattle.def = TempPlayer(index).PokemonBattle.def + (TempPlayer(index).PokemonBattle.def * (attackerstats.def * 0.15))
TempPlayer(index).PokemonBattle.spd = TempPlayer(index).PokemonBattle.spd + (TempPlayer(index).PokemonBattle.spd * (attackerstats.spd * 0.15))
TempPlayer(index).PokemonBattle.spatk = TempPlayer(index).PokemonBattle.spatk + (TempPlayer(index).PokemonBattle.spatk * (attackerstats.spatk * 0.15))
TempPlayer(index).PokemonBattle.spdef = TempPlayer(index).PokemonBattle.spdef + (TempPlayer(index).PokemonBattle.spdef * (attackerstats.spdef * 0.15))
End If



If message <> "" Then
Call SendBattleMessage(index, attackerName & " -> " & defenderName & " used " & usingMove & " - " & message, msgColor)
Else
Call SendBattleMessage(index, attackerName & " -> " & defenderName & " used " & usingMove, msgColor)
End If

If attackerStatus > STATUS_NOTHING Then
If attackerStatusNew > STATUS_NOTHING Then
attackerStatusNew = 0
End If
End If

If defenderStatus > STATUS_NOTHING Then
If defenderstatusnew > STATUS_NOTHING Then
defenderstatusnew = 0
End If
End If

Select Case attackerStatusNew
Case STATUS_PARALIZED
Call SendBattleMessage(index, attackerName & " is paralized.", Red)
Case STATUS_FREEZED
If attackerType1 = TYPE_ICE Or attackerType2 = TYPE_ICE Then
attackerStatusNew = 0
Else
Call SendBattleMessage(index, attackerName & " is frozen.", Red)
End If
Case STATUS_BURNED
If attackerType1 = TYPE_FIRE Or attackerType2 = TYPE_FIRE Then
attackerStatusNew = 0
Else
Call SendBattleMessage(index, attackerName & " is burned.", Red)
End If
Case STATUS_POISONED
If attackerType1 = TYPE_POISON Or attackerType2 = TYPE_POISON Or attackerType2 = TYPE_STEEL Or attackerType1 = TYPE_STEEL Then
attackerStatusNew = 0
Else
Call SendBattleMessage(index, attackerName & " is poisoned.", Red)
End If
Case STATUS_SLEEPING
'Check if is gym or NPC

Call SendBattleMessage(index, attackerName & " fell asleep.", Red)
Case STATUS_ATTRACTED
Call SendBattleMessage(index, attackerName & " is attracted.", Red)
Case STATUS_CONFUSED
Call SendBattleMessage(index, attackerName & " is confused.", Red)
Case STATUS_CURSED
Call SendBattleMessage(index, attackerName & " is cursed.", Red)
Case STATUS_FLINCHED
Call SendBattleMessage(index, attackerName & " is flinching.", Red)
Case STATUS_BADLYPOISONED
If attackerType1 = TYPE_POISON Or attackerType2 = TYPE_POISON Or attackerType2 = TYPE_STEEL Or attackerType1 = TYPE_STEEL Then
attackerStatusNew = 0
Else
Call SendBattleMessage(index, attackerName & " is badly poisoned.", Red)
End If
End Select
If bXY <= 0 Then
defenderstatusnew = 0
End If

Select Case defenderstatusnew
Case STATUS_PARALIZED
Call SendBattleMessage(index, defenderName & " is paralized.", Red)
Case STATUS_FREEZED
If defenderType1 = TYPE_ICE Or defenderType2 = TYPE_ICE Then
defenderstatusnew = 0
Else
Call SendBattleMessage(index, defenderName & " is frozen.", Red)
End If
Case STATUS_BURNED
If defenderType1 = TYPE_FIRE Or defenderType2 = TYPE_FIRE Then
defenderstatusnew = 0
Else
Call SendBattleMessage(index, defenderName & " is burned.", Red)
End If
Case STATUS_POISONED
If defenderType1 = TYPE_POISON Or defenderType2 = TYPE_POISON Or defenderType2 = TYPE_STEEL Or defenderType1 = TYPE_STEEL Then
defenderstatusnew = 0
Else
Call SendBattleMessage(index, defenderName & " is poisoned.", Red)
End If
Case STATUS_SLEEPING
If TempPlayer(index).inNPCBattle = True Then
Dim ai As Long
Dim xxx As Long
For ai = 1 To 6
If TempPlayer(index).NPCBattlePokemons(ai).status = STATUS_SLEEPING Then
xxx = xxx + 1
End If
Next
If xxx > 0 Then
Call SendBattleMessage(index, defenderName & " sleep failed!", Red)
defenderstatusnew = 0
Else
Call SendBattleMessage(index, defenderName & " fell asleep.", Red)
End If
Else
Call SendBattleMessage(index, defenderName & " fell asleep.", Red)
End If
Case STATUS_ATTRACTED
Call SendBattleMessage(index, defenderName & " is attracted.", Red)
Case STATUS_CONFUSED
Call SendBattleMessage(index, defenderName & " is confused.", Red)
Case STATUS_CURSED
Call SendBattleMessage(index, defenderName & " is cursed.", Red)
Case STATUS_FLINCHED
Call SendBattleMessage(index, defenderName & " is flinching.", Red)
Case STATUS_BADLYPOISONED
If defenderType1 = TYPE_POISON Or defenderType2 = TYPE_POISON Or defenderType2 = TYPE_STEEL Or defenderType1 = TYPE_STEEL Then
defenderstatusnew = 0
Else
Call SendBattleMessage(index, defenderName & " is badly poisoned.", Red)
End If
End Select


If attackerstats.atk <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " attack has been changed by " & attackerstats.atk, msgColor)
End If
If attackerstats.def <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " defense has been changed by " & attackerstats.def, msgColor)
End If
If attackerstats.spd <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " speed has been changed by " & attackerstats.spd, msgColor)
End If
If attackerstats.spatk <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " special attack has been changed by " & attackerstats.spatk, msgColor)
End If
If attackerstats.spdef <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " special defense has been changed by " & attackerstats.spdef, msgColor)
End If
If attackerstats.accuracy <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, attackerName & " accuracy has been changed by " & attackerstats.accuracy, msgColor)
End If
If defenderStats.atk <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " attack has been changed by " & defenderStats.atk, msgColor)
End If
If defenderStats.def <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " defense has been changed by " & defenderStats.def, msgColor)
End If
If defenderStats.spd <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " speed has been changed by " & defenderStats.spd, msgColor)
End If
If defenderStats.spatk <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " special attack has been changed by " & defenderStats.spatk, msgColor)
End If
If defenderStats.spdef <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " special defense has been changed by " & defenderStats.spdef, msgColor)
End If
If defenderStats.accuracy <> 0 Then
'ATK HAS CHANGED
Call SendBattleMessage(index, defenderName & " accuracy has been changed by " & defenderStats.accuracy, msgColor)
End If

'WRITE EVERYTHING TO MOVE DATA
'AFTER THIS IS WRITTEN USE MOVE SUB WILL USE IT TO DEAL DAMAGE SET STATS ETC.
TempPlayer(index).moveUsageTemp.attackerStatus = attackerStatusNew
TempPlayer(index).moveUsageTemp.attackerStatusRounds = attackerStatusNewTurns
TempPlayer(index).moveUsageTemp.customAttackUsage = customAttackUsage
TempPlayer(index).moveUsageTemp.damageInflict = damageInflict
TempPlayer(index).moveUsageTemp.defenderStatus = defenderstatusnew
TempPlayer(index).moveUsageTemp.defenderStatusRounds = defenderStatusNewTurns
TempPlayer(index).moveUsageTemp.effecrLast = effectLast
TempPlayer(index).moveUsageTemp.effectBegin = effectBegin
TempPlayer(index).moveUsageTemp.fleeBattle = fleeBattle
TempPlayer(index).moveUsageTemp.HPDamageModifier = HPDamageModifier
TempPlayer(index).moveUsageTemp.HPModifier = HPModifier
TempPlayer(index).moveUsageTemp.HPTotalModifier = HPTotalModifier
TempPlayer(index).moveUsageTemp.isCritical = isCritical
TempPlayer(index).moveUsageTemp.onlyOneMove = onlyOneMove
TempPlayer(index).moveUsageTemp.onlyOneMoveAdditionalEffect = onlyOneMoveAdditionalEffect
TempPlayer(index).moveUsageTemp.onlyOneMoveRounds = onlyOneMoveRounds
TempPlayer(index).moveUsageTemp.powerModifier = powerModifier
TempPlayer(index).moveUsageTemp.powerSet = powerSet
TempPlayer(index).moveUsageTemp.recoilHPCurrentModifier = recoilHPCurrentModifier
TempPlayer(index).moveUsageTemp.recoilHPDamageModifier = recoilHPDamageModifier
TempPlayer(index).moveUsageTemp.recoilHPTotalModifier = recoilHPTotalModifier
TempPlayer(index).moveUsageTemp.resetStats = resetStats
TempPlayer(index).moveUsageTemp.resetStatsTo = resetStatsTo
TempPlayer(index).moveUsageTemp.resetStatsMe = resetStatsMe
TempPlayer(index).moveUsageTemp.resetStatsMeTo = resetStatsMeTo
TempPlayer(index).moveUsageTemp.multiHit = hitTimes
'EVERYTHING ELSE IS DONE IN USE MOVE SUB



End Sub



Public Sub ClearMoveData(ByVal index As Long)
On Error Resume Next
Call ZeroMemory(ByVal VarPtr(TempPlayer(index).moveData), LenB(TempPlayer(index).moveData))
End Sub

Function getMoveFunctionCode(ByVal move As Long) As String
On Error Resume Next
  Dim fnCode As String
  Dim moveName As String
  moveName = Trim$(PokemonMove(move).Name)
  fnCode = GetVar(App.Path & "\Data\FunctionCodes.ini", "DATA", moveName)
  getMoveFunctionCode = fnCode
  
End Function
Function CheckPlayerDefeat(ByVal index As Long) As Boolean
On Error Resume Next
Dim i As Long
Dim n As Long

For i = 1 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 And player(index).PokemonInstance(i).hp > 0 Then
n = n + 1
End If
Next
If n > 0 Then
Exit Function
CheckPlayerDefeat = False
End If
Call SpawnPlayer(index)
Call HealPokemons(index)
CheckPlayerDefeat = True
End Function

Sub UseAnotherMove(ByVal index As Long, OnOpponent As Boolean, ByVal move As Long, ByVal slt As Long, Optional mslot As Long = 1)
On Error Resume Next
Call UseMove(index, OnOpponent, move, slt, mslot)
End Sub
Function GetMoveID(ByVal move As String) As Long
On Error Resume Next
Dim Num As Long
Dim moveTrim As String
moveTrim = Trim$(move)
Num = Val(GetVar(App.Path & "\Data\MoveNums.ini", "DATA", moveTrim))
GetMoveID = Num
End Function

Function GetTypeEffect(ByVal typeAtk As Byte, ByVal typeDef As Byte) As Double
On Error Resume Next
Select Case typeDef
Case TYPE_NONE
GetTypeEffect = 1
Case TYPE_NORMAL
GetTypeEffect = Types(typeAtk).NORMAL
Case TYPE_FIGHTING
GetTypeEffect = Types(typeAtk).FIGHT
Case TYPE_FLYING
GetTypeEffect = Types(typeAtk).FLYING
Case TYPE_POISON
GetTypeEffect = Types(typeAtk).POISON
Case TYPE_GROUND
GetTypeEffect = Types(typeAtk).GROUND
Case TYPE_ROCK
GetTypeEffect = Types(typeAtk).ROCK
Case TYPE_BUG
GetTypeEffect = Types(typeAtk).BUG
Case TYPE_GHOST
GetTypeEffect = Types(typeAtk).GHOST
Case TYPE_STEEL
GetTypeEffect = Types(typeAtk).STEEL
Case TYPE_FIRE
GetTypeEffect = Types(typeAtk).FIRE
Case TYPE_WATER
GetTypeEffect = Types(typeAtk).WATER
Case TYPE_GRASS
GetTypeEffect = Types(typeAtk).GRASS
Case TYPE_ELECTRIC
GetTypeEffect = Types(typeAtk).ELECTRIC
Case TYPE_PSYCHIC
GetTypeEffect = Types(typeAtk).PSYCHIC
Case TYPE_ICE
GetTypeEffect = Types(typeAtk).ICE
Case TYPE_DRAGON
GetTypeEffect = Types(typeAtk).DRAGON
Case TYPE_DARK
GetTypeEffect = Types(typeAtk).DARK
Case TYPE_FAIRY
GetTypeEffect = Types(typeAtk).FAIRY
End Select
End Function



Public Sub TryPokemonEvolution(ByVal index As Long, ByVal pokeSlot As Long)
On Error Resume Next
If Pokemon(player(index).PokemonInstance(pokeSlot).PokemonNumber).EvolvesTo > 0 Then
If Pokemon(player(index).PokemonInstance(pokeSlot).PokemonNumber).Evolution <= player(index).PokemonInstance(pokeSlot).level Then
If Trim$(Pokemon(player(index).PokemonInstance(pokeSlot).PokemonNumber).Stone) = "" Then
Call SendEvolve(index, pokeSlot, Pokemon(player(index).PokemonInstance(pokeSlot).PokemonNumber).EvolvesTo)
End If
End If
End If
End Sub

Public Sub EvolvePokemon(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
If Pokemon(player(index).PokemonInstance(slot).PokemonNumber).EvolvesTo > 0 Then
If Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Evolution <= player(index).PokemonInstance(slot).level Then
Else
Exit Sub
End If
Else
If Not TempPlayer(index).SpecialEvolveSlot > 0 Then
Exit Sub
End If
End If

Dim newPoke As Long
Dim oldPoke As Long
If TempPlayer(index).SpecialEvolveSlot > 0 Then
If TempPlayer(index).SpecialEvolveSlot = slot Then
newPoke = TempPlayer(index).SpecialEvolveTo
Else
newPoke = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).EvolvesTo
End If
Else
newPoke = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).EvolvesTo
End If
oldPoke = player(index).PokemonInstance(slot).PokemonNumber
player(index).PokemonInstance(slot).PokemonNumber = newPoke
player(index).PokemonInstance(slot).atk = player(index).PokemonInstance(slot).atk + (Pokemon(newPoke).atk - Pokemon(oldPoke).atk)
player(index).PokemonInstance(slot).def = player(index).PokemonInstance(slot).def + (Pokemon(newPoke).def - Pokemon(oldPoke).def)
player(index).PokemonInstance(slot).spd = player(index).PokemonInstance(slot).spd + (Pokemon(newPoke).spd - Pokemon(oldPoke).spd)
player(index).PokemonInstance(slot).spatk = player(index).PokemonInstance(slot).spatk + (Pokemon(newPoke).spatk - Pokemon(oldPoke).spatk)
player(index).PokemonInstance(slot).spdef = player(index).PokemonInstance(slot).spdef + (Pokemon(newPoke).spdef - Pokemon(oldPoke).spdef)
player(index).PokemonInstance(slot).MaxHp = player(index).PokemonInstance(slot).MaxHp + (Pokemon(newPoke).MaxHp - Pokemon(oldPoke).MaxHp)
player(index).PokemonInstance(slot).hp = player(index).PokemonInstance(slot).MaxHp
SendPlayerPokemon index
PlayerMsg index, "Your " & Trim$(Pokemon(oldPoke).Name) & " evolved to " & Trim$(Pokemon(newPoke).Name) & "!", BrightGreen
CheckForMoveLearn index, slot
If TempPlayer(index).SpecialEvolveItem > 0 Then
TakeItem index, TempPlayer(index).SpecialEvolveItem, 1
End If
TempPlayer(index).SpecialEvolveItem = 0
TempPlayer(index).SpecialEvolveSlot = 0
TempPlayer(index).SpecialEvolveTo = 0
End Sub

Public Sub CheckCustomLearnMove(ByVal index As Long, ByVal slot As Long, ByVal moveNum As Long)
On Error Resume Next
Dim n As Long
Dim i As Long
Dim realMoveNum As Long

Dim pNum As Long
pNum = player(index).PokemonInstance(slot).PokemonNumber
For i = 1 To 30
If Pokemon(pNum).moves(i) > 0 Then
n = n + 1
If n = moveNum Then
realMoveNum = i
End If
End If
Next
If player(index).PokemonInstance(slot).level = Pokemon(pNum).movesLV(realMoveNum) Then
For i = 1 To 4
If player(index).PokemonInstance(slot).moves(i).number = Pokemon(pNum).moves(realMoveNum) Then
Exit Sub
End If
Next
SendPlayerPokemon (index)
SendLearnMove index, slot, Pokemon(pNum).moves(realMoveNum)
TempPlayer(index).LearnMoveNumber = Pokemon(pNum).moves(realMoveNum)
TempPlayer(index).LearnMovePokemon = slot
TempPlayer(index).LearnMovePokemonName = Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name)
End If
End Sub


Sub NatureBonus(ByVal index As Long, ByVal slot As Long)
On Error Resume Next
PlayerMsg index, Trim$(Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Name) & " got nature boost!", BrightGreen
player(index).PokemonInstance(slot).MaxHp = player(index).PokemonInstance(slot).MaxHp + nature(player(index).PokemonInstance(slot).nature).AddHP
player(index).PokemonInstance(slot).atk = player(index).PokemonInstance(slot).atk + nature(player(index).PokemonInstance(slot).nature).AddAtk
player(index).PokemonInstance(slot).def = player(index).PokemonInstance(slot).def + nature(player(index).PokemonInstance(slot).nature).AddDef
player(index).PokemonInstance(slot).spatk = player(index).PokemonInstance(slot).spatk + nature(player(index).PokemonInstance(slot).nature).AddSpAtk
player(index).PokemonInstance(slot).spdef = player(index).PokemonInstance(slot).spdef + nature(player(index).PokemonInstance(slot).nature).AddSpDef
player(index).PokemonInstance(slot).spd = player(index).PokemonInstance(slot).spd + nature(player(index).PokemonInstance(slot).nature).AddSpd
SendPlayerPokemon (index)
End Sub



Sub WildDrop(ByVal index As Long)
On Error Resume Next
If TempPlayer(index).inNPCBattle = False Then
Dim rndNum As Long
Dim pokeName As String
pokeName = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name) & "C"
rndNum = Rand(1, 100)
Dim Chance As Long
Chance = Val(GetVar(App.Path & "\Data\Drop.ini", "DATA", pokeName))
If rndNum <= Chance Then
Call GiveItem(index, Val(GetVar(App.Path & "\Data\Drop.ini", "DATA", Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name))), 1)
End If
End If
End Sub


Sub LoadNPCBattle(ByVal index As Long, ByVal NPC As Long)
On Error Resume Next
TempPlayer(index).NPCBattle = NPC
Dim availablePokes As Long
availablePokes = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "DATA", "Pokemons"))
TempPlayer(index).NPCBattleSelectedPoke = Rand(1, availablePokes)
TempPlayer(index).NPCBattlePokesAvailable = availablePokes
Dim i As Long
For i = 1 To availablePokes
TempPlayer(index).NPCBattlePokemons(i).PokemonNumber = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "PokemonNumber"))
TempPlayer(index).NPCBattlePokemons(i).MaxHp = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "MaxHP"))
TempPlayer(index).NPCBattlePokemons(i).hp = TempPlayer(index).NPCBattlePokemons(i).MaxHp
TempPlayer(index).NPCBattlePokemons(i).level = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "level"))
TempPlayer(index).NPCBattlePokemons(i).atk = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "atk"))
TempPlayer(index).NPCBattlePokemons(i).def = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "def"))
TempPlayer(index).NPCBattlePokemons(i).spdef = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "spDef"))
TempPlayer(index).NPCBattlePokemons(i).spatk = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "spAtk"))
TempPlayer(index).NPCBattlePokemons(i).spd = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "Pokemon" & i, "spd"))
TempPlayer(index).NPCBattlePokemons(i).status = STATUS_NOTHING
TempPlayer(index).NPCBattlePokemons(i).turnsneed = 0
TempPlayer(index).NPCBattlePokemons(i).statusturn = 0
Next
End Sub

Sub StartNPCBattle(ByVal index As Long, ByVal customBackground As String)
On Error Resume Next
'This sub is called after NPC Battle info is loaded
'LOAD STARTING POKEMON
Dim slot As Long
Dim x As Long
If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If
If slot < 1 Then Exit Sub
'LOAD BATTLE POKE
TempPlayer(index).PokemonBattle = TempPlayer(index).NPCBattlePokemons(TempPlayer(index).NPCBattleSelectedPoke)
'END
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
For x = 1 To 6
player(index).PokemonInstance(x).batk = player(index).PokemonInstance(x).atk
player(index).PokemonInstance(x).bdef = player(index).PokemonInstance(x).def
player(index).PokemonInstance(x).bspd = player(index).PokemonInstance(x).spd
player(index).PokemonInstance(x).bspatk = player(index).PokemonInstance(x).spatk
player(index).PokemonInstance(x).bspdef = player(index).PokemonInstance(x).spdef
Next
TempPlayer(index).BattleCurrentTurn = 1
TempPlayer(index).inNPCBattle = True
SendNpcBattle index, slot, YES, YES, "GymBattle1.mp3", customBackground
End Sub

Sub NPCSwitchPoke(ByVal index As Long, ByVal newPoke As Long, ByVal slot As Long)
On Error Resume Next
TempPlayer(index).NPCBattlePokemons(TempPlayer(index).NPCBattleSelectedPoke) = TempPlayer(index).PokemonBattle
TempPlayer(index).PokemonBattle = TempPlayer(index).NPCBattlePokemons(newPoke)
TempPlayer(index).NPCBattleSelectedPoke = newPoke
'END
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
TempPlayer(index).BattleCurrentTurn = 1
SendNpcBattle index, slot, YES
End Sub

Sub CheckNpcWinData(ByVal index As Long)
On Error Resume Next
Dim NPC As Long
Dim badge As Long
NPC = TempPlayer(index).NPCBattle
If GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "DATA", "isGym") = "YES" Then
badge = Val(GetVar(App.Path & "\Data\NPCBattles\NPCBattle" & NPC & ".ini", "DATA", "Badge"))
If player(index).Bedages(badge) <> YES Then
player(index).Bedages(badge) = YES
GlobalMsg "[GYM]" & Trim$(player(index).Name) & " has won Gym " & badge & " badge!", BrightGreen
GiveItem index, 1, 1000
End If
End If
End Sub

Public Function GetWildAI(ByVal index As Long, ByVal slot As Long) As Long
On Error Resume Next
Dim mineType1 As Long
Dim mineType2 As Long
Dim mineType1Text As String
Dim mineType2Text As String
Dim pokeName As String
Dim pokeLevel As Long
pokeLevel = TempPlayer(index).PokemonBattle.level
pokeName = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)
mineType1 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type
mineType2 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type2
If mineType1 <= 0 And mineType2 > 0 Then
mineType1 = mineType2
mineType2 = 0
End If
mineType1Text = TypeToText(mineType1)
mineType2Text = TypeToText(mineType2)
'Get AI
If TempPlayer(index).PokemonBattle.FirstMove = NO Then
If Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", "FirstMove")) <> "" Then
TempPlayer(index).PokemonBattle.FirstMove = YES
GetWildAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", "FirstMove")))
Exit Function
End If
End If

If mineType1 > 0 And mineType2 > 0 Then
If Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType1Text & "|" & mineType2Text)) <> "" Then
GetWildAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType1Text & "|" & mineType2Text)))
Exit Function
End If
'Try opposite way
If Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType2Text & "|" & mineType1Text)) <> "" Then
GetWildAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType2Text & "|" & mineType1Text)))
Exit Function
End If
End If

If Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType1Text)) <> "" Then
GetWildAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType1Text)))
Exit Function
End If


If Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType2Text)) <> "" Then
GetWildAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\WildAI\" & pokeName & "LV" & pokeLevel & ".ini", "DATA", mineType2Text)))
Exit Function
End If


'No ai found
GetWildAI = 0
End Function


Public Function GetNPCAI(ByVal index As Long, ByVal slot As Long) As Long
On Error Resume Next
Dim mineType1 As Long
Dim mineType2 As Long
Dim mineType1Text As String
Dim mineType2Text As String
Dim pokeName As String
Dim pokeLevel As Long
Dim NpcNum As Long
Dim npcPoke As Long
NpcNum = TempPlayer(index).NPCBattle
npcPoke = TempPlayer(index).NPCBattleSelectedPoke
pokeLevel = TempPlayer(index).PokemonBattle.level
pokeName = Trim$(Pokemon(TempPlayer(index).PokemonBattle.PokemonNumber).Name)
mineType1 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type
mineType2 = Pokemon(player(index).PokemonInstance(slot).PokemonNumber).Type2
If mineType1 <= 0 And mineType2 > 0 Then
mineType1 = mineType2
mineType2 = 0
End If
mineType1Text = TypeToText(mineType1)
mineType2Text = TypeToText(mineType2)
'Get AI
If TempPlayer(index).PokemonBattle.FirstMove = NO Then
If Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, "FirstMove")) <> "" Then
TempPlayer(index).PokemonBattle.FirstMove = YES
GetNPCAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, "FirstMove")))
Exit Function
End If
End If

If mineType1 > 0 And mineType2 > 0 Then
If Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType1Text & "|" & mineType2Text)) <> "" Then
GetNPCAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType1Text & "|" & mineType2Text)))
Exit Function
End If
'Try opposite way
If Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType2Text & "|" & mineType1Text)) <> "" Then
GetNPCAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType2Text & "|" & mineType1Text)))
Exit Function
End If
End If

If Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType1Text)) <> "" Then
GetNPCAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType1Text)))
Exit Function
End If


If Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType2Text)) <> "" Then
GetNPCAI = GetMoveID(Trim$(GetVar(App.Path & "\Data\NpcAI\NPC" & NpcNum & ".ini", "Pokemon" & npcPoke, mineType2Text)))
Exit Function
End If


'No ai found
GetNPCAI = 0
End Function













Public Function TypeToText(ByVal typeNum As Long) As String
On Error Resume Next
Select Case typeNum
Case TYPE_NONE
TypeToText = "None"
Case TYPE_NORMAL
TypeToText = "Normal"
Case TYPE_FIGHTING
TypeToText = "Fighting"
Case TYPE_FLYING
TypeToText = "Flying"
Case TYPE_POISON
TypeToText = "Poison"
Case TYPE_GROUND
TypeToText = "Ground"
Case TYPE_ROCK
TypeToText = "Rock"
Case TYPE_BUG
TypeToText = "Bug"
Case TYPE_GHOST
TypeToText = "Ghost"
Case TYPE_STEEL
TypeToText = "Steel"
Case TYPE_FIRE
TypeToText = "Fire"
Case TYPE_WATER
TypeToText = "Water"
Case TYPE_GRASS
TypeToText = "Grass"
Case TYPE_ELECTRIC
TypeToText = "Electric"
Case TYPE_PSYCHIC
TypeToText = "Psychic"
Case TYPE_ICE
TypeToText = "Ice"
Case TYPE_DRAGON
TypeToText = "Dragon"
Case TYPE_DARK
TypeToText = "Dark"
Case TYPE_FAIRY
TypeToText = "Fairy"
End Select
End Function

Public Sub initFishBattle(ByVal index As Long)
On Error Resume Next
ResetBattlePokemon (index)
Dim i As Long
Dim x As Long
Dim wildpoke As Long
Dim slot As Long
Dim frmlvl As Long
Dim tolvl As Long
Dim cstm As Long
Dim slt As Long

If TempPlayer(index).PokemonBattle.PokemonNumber > 0 Then Exit Sub

If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If

If slot < 1 Then Exit Sub


'Set wild pokemon
For i = 1 To GetMapFishPokes(player(index).map)
If SpawnChance(GetMapFishPokeChance(player(index).map, 1)) = True Then
wildpoke = GetMapFishPokeNumber(player(index).map, i)
frmlvl = GetMapFishPokeFrom(player(index).map, i)
tolvl = GetMapFishPokeTo(player(index).map, i)
cstm = NO
slt = i
Exit For
End If
Next



If wildpoke < 1 Or wildpoke > 721 Then Exit Sub 'No battle if there is not pokemon to spawn

'If there is pokemon then we are going to set BattlePokemon ready!

ResetBattlePokemon (index)
TempPlayer(index).PokemonBattle.PokemonNumber = wildpoke
TempPlayer(index).PokemonBattle.level = Rand(frmlvl, tolvl)
TempPlayer(index).PokemonBattle.MapSlot = slt
TempPlayer(index).PokemonBattle.nature = Rand(1, MAX_NATURES)
TempPlayer(index).PokemonBattle.status = STATUS_NOTHING
TempPlayer(index).PokemonBattle.turnsneed = 0
TempPlayer(index).PokemonBattle.statusturn = 0
Dim rndNum As Long
rndNum = Rand(1, 100000)
If GetPlayerAccess(index) >= ADMIN_DEVELOPER Then
'PlayerMsg index, "Shiny num: " & rndnum, Yellow
End If
If rndNum = 23 Then
TempPlayer(index).PokemonBattle.isShiny = YES ' for now
GlobalMsg "(SHINY!) " & Trim$(player(index).Name) & " encountered lvl." & TempPlayer(index).PokemonBattle.level & " shiny " & Trim$(Pokemon(wildpoke).Name), Pink
End If

'MoveThing for now
TempPlayer(index).PokemonBattle.moves(1).number = Pokemon(wildpoke).moves(1)
Select Case cstm
Case NO
TempPlayer(index).PokemonBattle.atk = CalculateStat(Pokemon(wildpoke).atk, STAT_ATK)
TempPlayer(index).PokemonBattle.def = CalculateStat(Pokemon(wildpoke).def, STAT_DEF)
TempPlayer(index).PokemonBattle.spatk = CalculateStat(Pokemon(wildpoke).spatk, STAT_SPATK)
TempPlayer(index).PokemonBattle.spdef = CalculateStat(Pokemon(wildpoke).spdef, STAT_SPDEF)
TempPlayer(index).PokemonBattle.spd = CalculateStat(Pokemon(wildpoke).spd, STAT_SPEED)
TempPlayer(index).PokemonBattle.MaxHp = CalculateStat(Pokemon(wildpoke).MaxHp, STAT_HP)
If TempPlayer(index).PokemonBattle.level > 1 Then
Dim availableTP As Long
availableTP = TempPlayer(index).PokemonBattle.level * 3 - 3
Do While availableTP > 0
Dim stattoadd As Long
stattoadd = Rand(1, 6)
Select Case stattoadd
Case STAT_ATK
TempPlayer(index).PokemonBattle.atk = TempPlayer(index).PokemonBattle.atk + 1
Case STAT_DEF
TempPlayer(index).PokemonBattle.def = TempPlayer(index).PokemonBattle.def + 1
Case STAT_SPATK
TempPlayer(index).PokemonBattle.spatk = TempPlayer(index).PokemonBattle.spatk + 1
Case STAT_SPDEF
TempPlayer(index).PokemonBattle.spdef = TempPlayer(index).PokemonBattle.spdef + 1
Case STAT_SPEED
TempPlayer(index).PokemonBattle.spd = TempPlayer(index).PokemonBattle.spd + 1
Case STAT_HP
TempPlayer(index).PokemonBattle.MaxHp = TempPlayer(index).PokemonBattle.MaxHp + 2
End Select
availableTP = availableTP - 1
Loop
End If
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.MaxHp

End Select

For x = 1 To 6
player(index).PokemonInstance(x).batk = player(index).PokemonInstance(x).atk
player(index).PokemonInstance(x).bdef = player(index).PokemonInstance(x).def
player(index).PokemonInstance(x).bspd = player(index).PokemonInstance(x).spd
player(index).PokemonInstance(x).bspatk = player(index).PokemonInstance(x).spatk
player(index).PokemonInstance(x).bspdef = player(index).PokemonInstance(x).spdef
Next



'Set turn (My Speed>Enemy Speed = MyTurn)
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
TempPlayer(index).BattleCurrentTurn = 1
SendNpcBattle index, slot
Call SendActionMsg(GetPlayerMap(index), "Encounter: " & Trim$(Pokemon(wildpoke).Name), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32)

End Sub




Function GetMapFishPokes(ByVal map As Long)
On Error Resume Next
GetMapFishPokes = Val(GetVar(App.Path & "\Data\FishSpawns\" & map & ".ini", "DATA", "Spawns"))
End Function
Function GetMapFishPokeChance(ByVal map As Long, ByVal poke As Long)
On Error Resume Next
GetMapFishPokeChance = Val(GetVar(App.Path & "\Data\FishSpawns\" & map & ".ini", "DATA", poke & "Chance"))
End Function
Function GetMapFishPokeNumber(ByVal map As Long, ByVal poke As Long)
On Error Resume Next
GetMapFishPokeNumber = Val(GetVar(App.Path & "\Data\FishSpawns\" & map & ".ini", "DATA", poke & "Number"))
End Function

Function GetMapFishPokeFrom(ByVal map As Long, ByVal poke As Long)
On Error Resume Next
GetMapFishPokeFrom = Val(GetVar(App.Path & "\Data\FishSpawns\" & map & ".ini", "DATA", poke & "From"))
End Function
Function GetMapFishPokeTo(ByVal map As Long, ByVal poke As Long)
On Error Resume Next
GetMapFishPokeTo = Val(GetVar(App.Path & "\Data\FishSpawns\" & map & ".ini", "DATA", poke & "To"))
End Function





Function GetMoveName(ByVal moveNum As Long) As String
If moveNum < 1 Or moveNum > MAX_MOVES Then Exit Function
GetMoveName = Trim$(PokemonMove(moveNum).Name)
End Function


Public Sub initHoneyBattle(ByVal index As Long)
On Error Resume Next
ResetBattlePokemon (index)
Dim i As Long
Dim x As Long
Dim wildpoke As Long
Dim slot As Long
Dim frmlvl As Long
Dim tolvl As Long
Dim cstm As Long
Dim slt As Long

If TempPlayer(index).PokemonBattle.PokemonNumber > 0 Then Exit Sub

If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If

If slot < 1 Then Exit Sub


'Set wild pokemon
Dim pokes As Long
pokes = GetHoneyPokes(index)
Dim selPoke As Long
selPoke = Rand(1, pokes)
wildpoke = GetHoneyPoke(index, selPoke)
frmlvl = GetHoneyPokeLevel(index, selPoke)
tolvl = GetHoneyPokeLevel(index, selPoke)



If wildpoke < 1 Or wildpoke > 721 Then Exit Sub 'No battle if there is not pokemon to spawn

'If there is pokemon then we are going to set BattlePokemon ready!

ResetBattlePokemon (index)
TempPlayer(index).PokemonBattle.PokemonNumber = wildpoke
TempPlayer(index).PokemonBattle.level = Rand(frmlvl, tolvl)
TempPlayer(index).PokemonBattle.MapSlot = slt
TempPlayer(index).PokemonBattle.nature = Rand(1, MAX_NATURES)
TempPlayer(index).PokemonBattle.status = STATUS_NOTHING
TempPlayer(index).PokemonBattle.turnsneed = 0
TempPlayer(index).PokemonBattle.statusturn = 0
Dim rndNum As Long
rndNum = Rand(1, 100000)
If GetPlayerAccess(index) >= ADMIN_DEVELOPER Then
'PlayerMsg index, "Shiny num: " & rndnum, Yellow
End If
If rndNum = 23 Then
TempPlayer(index).PokemonBattle.isShiny = YES ' for now
GlobalMsg "(SHINY!) " & Trim$(player(index).Name) & " encountered lvl." & TempPlayer(index).PokemonBattle.level & " shiny " & Trim$(Pokemon(wildpoke).Name), Pink
End If

'MoveThing for now
TempPlayer(index).PokemonBattle.moves(1).number = Pokemon(wildpoke).moves(1)
Select Case cstm
Case NO
TempPlayer(index).PokemonBattle.atk = CalculateStat(Pokemon(wildpoke).atk, STAT_ATK)
TempPlayer(index).PokemonBattle.def = CalculateStat(Pokemon(wildpoke).def, STAT_DEF)
TempPlayer(index).PokemonBattle.spatk = CalculateStat(Pokemon(wildpoke).spatk, STAT_SPATK)
TempPlayer(index).PokemonBattle.spdef = CalculateStat(Pokemon(wildpoke).spdef, STAT_SPDEF)
TempPlayer(index).PokemonBattle.spd = CalculateStat(Pokemon(wildpoke).spd, STAT_SPEED)
TempPlayer(index).PokemonBattle.MaxHp = CalculateStat(Pokemon(wildpoke).MaxHp, STAT_HP)
If TempPlayer(index).PokemonBattle.level > 1 Then
Dim availableTP As Long
availableTP = TempPlayer(index).PokemonBattle.level * 3 - 3
Do While availableTP > 0
Dim stattoadd As Long
stattoadd = Rand(1, 6)
Select Case stattoadd
Case STAT_ATK
TempPlayer(index).PokemonBattle.atk = TempPlayer(index).PokemonBattle.atk + 1
Case STAT_DEF
TempPlayer(index).PokemonBattle.def = TempPlayer(index).PokemonBattle.def + 1
Case STAT_SPATK
TempPlayer(index).PokemonBattle.spatk = TempPlayer(index).PokemonBattle.spatk + 1
Case STAT_SPDEF
TempPlayer(index).PokemonBattle.spdef = TempPlayer(index).PokemonBattle.spdef + 1
Case STAT_SPEED
TempPlayer(index).PokemonBattle.spd = TempPlayer(index).PokemonBattle.spd + 1
Case STAT_HP
TempPlayer(index).PokemonBattle.MaxHp = TempPlayer(index).PokemonBattle.MaxHp + 2
End Select
availableTP = availableTP - 1
Loop
End If
TempPlayer(index).PokemonBattle.hp = TempPlayer(index).PokemonBattle.MaxHp

End Select

For x = 1 To 6
player(index).PokemonInstance(x).batk = player(index).PokemonInstance(x).atk
player(index).PokemonInstance(x).bdef = player(index).PokemonInstance(x).def
player(index).PokemonInstance(x).bspd = player(index).PokemonInstance(x).spd
player(index).PokemonInstance(x).bspatk = player(index).PokemonInstance(x).spatk
player(index).PokemonInstance(x).bspdef = player(index).PokemonInstance(x).spdef
Next



'Set turn (My Speed>Enemy Speed = MyTurn)
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
TempPlayer(index).BattleCurrentTurn = 1
SendNpcBattle index, slot
Call SendActionMsg(GetPlayerMap(index), "Encounter: " & Trim$(Pokemon(wildpoke).Name), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32)

End Sub

Function GetHoneyPokes(ByVal index As Long)
On Error Resume Next
GetHoneyPokes = Val(GetVar(App.Path & "\Data\Honey.ini", GetPlayerMap(index), "POKES"))
End Function

Function GetHoneyPoke(ByVal index As Long, ByVal Num As Long)
On Error Resume Next
Dim numStr As String
numStr = Num
GetHoneyPoke = Val(GetVar(App.Path & "\Data\Honey.ini", GetPlayerMap(index), "NUMBER" & numStr))
End Function

Function GetHoneyPokeLevel(ByVal index As Long, ByVal Num As Long)
On Error Resume Next
Dim numStr As String
numStr = Num
GetHoneyPokeLevel = Val(GetVar(App.Path & "\Data\Honey.ini", GetPlayerMap(index), "LEVEL" & numStr))
End Function

Sub StartPVPBattle(ByVal index As Long)
On Error Resume Next
'This sub is called after NPC Battle info is loaded
'LOAD STARTING POKEMON
Dim enemyIndex As Long
enemyIndex = FindPlayer(Trim$(TempPlayer(index).PVPEnemy))
Dim slot As Long
Dim x As Long
If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If
If slot < 1 Then Exit Sub
'LOAD BATTLE POKE

'END
If player(index).PokemonInstance(slot).spd > TempPlayer(index).PokemonBattle.spd Then 'If my speed is bigger than enemys then I attack first
TempPlayer(index).BattleTurn = True
Else
If player(index).PokemonInstance(slot).spd = TempPlayer(index).PokemonBattle.spd Then ' if its equal as enemys then 50% to be mine else its enemys
If SpawnChance(2) = True Then
TempPlayer(index).BattleTurn = True
Else
TempPlayer(index).BattleTurn = False
End If
Else
TempPlayer(index).BattleTurn = False 'If enemys is bigger then its his turn
End If
End If
For x = 1 To 6
player(index).PokemonInstance(x).batk = player(index).PokemonInstance(x).atk
player(index).PokemonInstance(x).bdef = player(index).PokemonInstance(x).def
player(index).PokemonInstance(x).bspd = player(index).PokemonInstance(x).spd
player(index).PokemonInstance(x).bspatk = player(index).PokemonInstance(x).spatk
player(index).PokemonInstance(x).bspdef = player(index).PokemonInstance(x).spdef
Next
If player(index).PokemonInstance(1).hp > 0 And player(index).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(index).PokemonInstance(i).PokemonNumber > 0 Then
If player(index).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(index, i)
slot = 1
Exit For
Else
End If
End If
Next
End If


'OTHER PLAYER
slot = 0
x = 0
If player(enemyIndex).PokemonInstance(1).hp > 0 And player(enemyIndex).PokemonInstance(1).PokemonNumber > 0 Then
slot = 1
Else
For i = 2 To 6
If player(enemyIndex).PokemonInstance(i).PokemonNumber > 0 Then
If player(enemyIndex).PokemonInstance(i).hp > 0 Then
Call SetAsLeader(enemyIndex, i)
slot = 1
Exit For
Else
End If
End If
Next
End If
If slot < 1 Then Exit Sub
'LOAD BATTLE POKE

'END
TempPlayer(enemyIndex).BattleTurn = Not TempPlayer(index).BattleTurn
For x = 1 To 6
player(enemyIndex).PokemonInstance(x).batk = player(enemyIndex).PokemonInstance(x).atk
player(enemyIndex).PokemonInstance(x).bdef = player(enemyIndex).PokemonInstance(x).def
player(enemyIndex).PokemonInstance(x).bspd = player(enemyIndex).PokemonInstance(x).spd
player(enemyIndex).PokemonInstance(x).bspatk = player(enemyIndex).PokemonInstance(x).spatk
player(enemyIndex).PokemonInstance(x).bspdef = player(enemyIndex).PokemonInstance(x).spdef
Next

SetPokemonBattlePVP index, 1, enemyIndex
SetPokemonBattlePVP enemyIndex, 1, index

TempPlayer(index).BattleCurrentTurn = 1

SendNpcBattle index, slot, YES

TempPlayer(enemyIndex).BattleCurrentTurn = 1

SendNpcBattle enemyIndex, slot, YES
End Sub

Sub SetPokemonBattlePVP(ByVal index As Long, ByVal enemySlot As Long, ByVal enemyIndex As Long)
On Error Resume Next
TempPlayer(index).PokemonBattle.PokemonNumber = player(enemyIndex).PokemonInstance(enemySlot).PokemonNumber
TempPlayer(index).PokemonBattle.atk = player(enemyIndex).PokemonInstance(enemySlot).batk
TempPlayer(index).PokemonBattle.def = player(enemyIndex).PokemonInstance(enemySlot).bdef
TempPlayer(index).PokemonBattle.spatk = player(enemyIndex).PokemonInstance(enemySlot).bspatk
TempPlayer(index).PokemonBattle.spdef = player(enemyIndex).PokemonInstance(enemySlot).bspdef
TempPlayer(index).PokemonBattle.spd = player(enemyIndex).PokemonInstance(enemySlot).bspd
TempPlayer(index).PokemonBattle.hp = player(enemyIndex).PokemonInstance(enemySlot).hp
TempPlayer(index).PokemonBattle.MaxHp = player(enemyIndex).PokemonInstance(enemySlot).MaxHp
TempPlayer(index).PokemonBattle.level = player(enemyIndex).PokemonInstance(enemySlot).level

End Sub
Sub PVPProcessRound(ByVal index As Long)
On Error Resume Next
Dim enemyIndex As Long
Dim mainSlot As Long
Dim enemySlot As Long
Dim move As Long
enemyIndex = FindPlayer(Trim$(TempPlayer(index).PVPEnemy))
mainSlot = TempPlayer(index).PVPSlot
enemySlot = TempPlayer(enemyIndex).PVPSlot

If TempPlayer(index).PVPCommandUsed = PVP_SWITCH Then
If TempPlayer(index).WaitingForSwitch = YES Then
SetAsLeader index, TempPlayer(index).PVPCommandNum
TempPlayer(index).PVPSlot = 1
UpdatePVPEnemyPokemonBattle index, enemyIndex, enemySlot, TempPlayer(index).PVPSlot

TempPlayer(enemyIndex).PVPHasUsed = False
TempPlayer(index).PVPHasUsed = False
TempPlayer(index).WaitingForSwitch = NO

SendBattleUpdate index, TempPlayer(index).PVPSlot, 0, YES
SendPlayerPokemon index
SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendPlayerPokemon enemyIndex
Exit Sub
Else
TempPlayer(index).PVPSlot = TempPlayer(index).PVPCommandNum
mainSlot = TempPlayer(index).PVPSlot
UpdatePVPEnemyPokemonBattle index, enemyIndex, enemySlot, TempPlayer(index).PVPSlot
TempPlayer(enemyIndex).PVPTurnAdvantage = True

TempPlayer(enemyIndex).PVPHasUsed = False
TempPlayer(index).PVPHasUsed = False
End If
End If



'ENEMYSWITHC
If TempPlayer(enemyIndex).PVPCommandUsed = PVP_SWITCH Then
If TempPlayer(enemyIndex).WaitingForSwitch = YES Then
SetAsLeader enemyIndex, TempPlayer(enemyIndex).PVPCommandNum
TempPlayer(enemyIndex).PVPSlot = 1
UpdatePVPEnemyPokemonBattle index, enemyIndex, TempPlayer(index).PVPSlot, mainSlot

TempPlayer(enemyIndex).PVPHasUsed = False
TempPlayer(index).PVPHasUsed = False
TempPlayer(enemyIndex).WaitingForSwitch = NO

SendBattleUpdate index, mainSlot, 0, YES
SendPlayerPokemon index
SendBattleUpdate enemyIndex, TempPlayer(index).PVPSlot, 0, YES
SendPlayerPokemon enemyIndex
Exit Sub

Else
TempPlayer(enemyIndex).PVPSlot = TempPlayer(enemyIndex).PVPCommandNum
enemySlot = TempPlayer(enemyIndex).PVPSlot
UpdatePVPMainPokemonBattle index, enemyIndex, TempPlayer(enemyIndex).PVPSlot, mainSlot
TempPlayer(index).PVPTurnAdvantage = True
TempPlayer(enemyIndex).PVPHasUsed = False
TempPlayer(index).PVPHasUsed = False

End If
End If

If TempPlayer(index).PVPCommandUsed = PVP_SWITCH Or TempPlayer(enemyIndex).PVPCommandUsed = PVP_SWITCH Then
SendBattleUpdate index, mainSlot, 0, YES
SendPlayerPokemon index
SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendPlayerPokemon enemyIndex
Exit Sub

End If



'OTHER!

Select Case TempPlayer(index).PVPCommandUsed

Case PVP_MOVE
CheckTurn index, mainSlot

If TempPlayer(index).PVPTurnAdvantage = True Then
TempPlayer(index).BattleTurn = True
TempPlayer(index).PVPTurnAdvantage = False
End If
If TempPlayer(enemyIndex).PVPTurnAdvantage = True Then
TempPlayer(enemyIndex).BattleTurn = True
TempPlayer(enemyIndex).PVPTurnAdvantage = False
End If

If TempPlayer(index).BattleTurn = True Then
'Index attacks first
If isPlayerDefeated(index, TempPlayer(index).PVPSlot) Then
If DoesPlayerHavePokemons(index) Then
SendBattleUpdate index, mainSlot, 0, YES
SendBattleUpdate enemyIndex, enemySlot
TempPlayer(index).WaitingForSwitch = YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

TempPlayer(index).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(index, enemyIndex) 'Wiped out
Exit Sub
End If
End If


'WE PASSED CONTINUE ATTACKING
PlayerAttackWild index, TempPlayer(index).PVPCommandNum, mainSlot
UpdatePVPEnemyPokemon index, enemyIndex, enemySlot, mainSlot
If IsWildDefeated(index) Then
If DoesPlayerHavePokemons(enemyIndex) Then
SendBattleUpdate index, mainSlot
SendBattleUpdate enemyIndex, enemySlot, 0, YES
TempPlayer(enemyIndex).WaitingForSwitch = YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

TempPlayer(enemyIndex).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(index, index) 'Wiped out
Exit Sub
End If
End If


'WE PASSED THIS ONE TOO
If TempPlayer(enemyIndex).PVPCommandUsed = PVP_MOVE Then
PlayerAttackWild enemyIndex, TempPlayer(enemyIndex).PVPCommandNum, enemySlot, False
UpdatePVPMainPokemon index, enemyIndex, enemySlot, mainSlot
If IsWildDefeated(enemyIndex) Then
If DoesPlayerHavePokemons(index) Then
SendBattleUpdate enemyIndex, enemySlot
SendBattleUpdate index, mainSlot, 0, YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex
TempPlayer(index).WaitingForSwitch = YES
TempPlayer(index).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(enemyIndex, enemyIndex) 'Wiped out
Exit Sub
End If
End If
End If
SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendBattleUpdate index, mainSlot, 0, YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

 TempPlayer(index).PVPHasUsed = False
 TempPlayer(enemyIndex).PVPHasUsed = False
 


Else
'enemyIndex attacks
If TempPlayer(index).PVPCommandUsed = PVP_MOVE Then
If isPlayerDefeated(enemyIndex, TempPlayer(enemyIndex).PVPSlot) Then
If DoesPlayerHavePokemons(enemyIndex) Then
SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendBattleUpdate index, mainSlot
TempPlayer(enemyIndex).WaitingForSwitch = YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

TempPlayer(enemyIndex).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(enemyIndex, index) 'Wiped out
Exit Sub
End If
End If


'WE PASSED CONTINUE ATTACKING
PlayerAttackWild enemyIndex, TempPlayer(enemyIndex).PVPCommandNum, enemySlot
UpdatePVPMainPokemon index, enemyIndex, enemySlot, mainSlot
If IsWildDefeated(enemyIndex) Then
If DoesPlayerHavePokemons(index) Then
SendBattleUpdate index, mainSlot, 0, YES
SendBattleUpdate enemyIndex, enemySlot
TempPlayer(index).WaitingForSwitch = YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

TempPlayer(index).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(enemyIndex, enemyIndex) 'Wiped out
Exit Sub
End If
End If


'WE PASSED THIS ONE TOO
PlayerAttackWild index, TempPlayer(index).PVPCommandNum, mainSlot, False
UpdatePVPEnemyPokemon index, enemyIndex, enemySlot, mainSlot
If IsWildDefeated(index) Then
If DoesPlayerHavePokemons(enemyIndex) Then
SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendBattleUpdate index, mainSlot, 0, 0
TempPlayer(enemyIndex).WaitingForSwitch = YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

TempPlayer(enemyIndex).PVPHasUsed = False
Exit Sub
Else
Call EndPVP(index, index) 'Wiped out
Exit Sub
End If
End If

SendBattleUpdate enemyIndex, enemySlot, 0, YES
SendBattleUpdate index, mainSlot, 0, YES
SendPlayerPokemon index
SendPlayerPokemon enemyIndex

 TempPlayer(index).PVPHasUsed = False
 TempPlayer(enemyIndex).PVPHasUsed = False
 

End If
End If


End Select

End Sub


Function PVPPlayerDefeated(ByVal index As Long, ByVal slot As Long)

End Function

Function DoesPlayerHavePokemons(ByVal index As Long) As Boolean
On Error Resume Next
Dim n As Long
For n = 1 To 6
If player(index).PokemonInstance(n).PokemonNumber > 0 And player(index).PokemonInstance(n).hp > 0 Then
DoesPlayerHavePokemons = True
Exit Function



End If
Next
End Function
Sub EndPVP(ByVal index As Long, ByVal winner As Long)
On Error Resume Next
Dim enemyIndex As Long
enemyIndex = FindPlayer(Trim$(TempPlayer(index).PVPEnemy))
SendPVPCommand index, "NOTPVP"
SendPVPCommand enemyIndex, "NOTPVP"
TempPlayer(index).isInPVP = False
TempPlayer(index).PVPCommandNum = 0
TempPlayer(index).PVPCommandUsed = 0
TempPlayer(index).PVPEnemy = ""
TempPlayer(index).PVPHasUsed = False
TempPlayer(index).PVPSlot = 0

TempPlayer(enemyIndex).PVPCommandNum = 0
TempPlayer(enemyIndex).PVPCommandUsed = 0
TempPlayer(enemyIndex).PVPEnemy = ""
TempPlayer(enemyIndex).PVPHasUsed = False
TempPlayer(enemyIndex).PVPSlot = 0

If winner = enemyIndex Then
SendBattleInfo enemyIndex, 0, YES, 0
Call SendBattleInfo(index, 0, BATTLE_NO, 0)
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slot
Call SpawnPlayer(index)
Call HealPokemons(index)
If CanAddRankPoint(index, enemyIndex) Then
AddRankedPoint enemyIndex
RemoveRankedPoint index
End If
Else
Call SendBattleInfo(enemyIndex, 0, BATTLE_NO, 0)
ResetBattlePokemon (enemyIndex)
SendPlayerPokemon enemyIndex
SendBattleUpdate enemyIndex, slot
Call SpawnPlayer(enemyIndex)
Call HealPokemons(enemyIndex)
SendBattleInfo index, 0, YES, 0
If CanAddRankPoint(index, enemyIndex) Then
AddRankedPoint index
RemoveRankedPoint enemyIndex
End If
End If
'
TempPlayer(index).isInPVP = False
TempPlayer(index).BattleType = 0
TempPlayer(index).BattleCurrentTurn = 0
ResetBattlePokemon (index)
SendPlayerPokemon index
SendBattleUpdate index, slot
'
TempPlayer(enemyIndex).isInPVP = False

TempPlayer(enemyIndex).BattleType = 0
TempPlayer(enemyIndex).BattleCurrentTurn = 0
ResetBattlePokemon (enemyIndex)
SendPlayerPokemon enemyIndex
SendBattleUpdate enemyIndex, slot
MapMsg player(index).map, Trim$(player(winner).Name) & " has won a battle!", White

End Sub

Sub UpdatePVPEnemyPokemon(ByVal index As Long, ByVal enemyIndex As Long, ByVal enemySlot As Long, ByVal mainSlot As Long)
On Error Resume Next
player(enemyIndex).PokemonInstance(enemySlot).batk = TempPlayer(index).PokemonBattle.atk
player(enemyIndex).PokemonInstance(enemySlot).bdef = TempPlayer(index).PokemonBattle.def
player(enemyIndex).PokemonInstance(enemySlot).bspatk = TempPlayer(index).PokemonBattle.spatk
player(enemyIndex).PokemonInstance(enemySlot).bspdef = TempPlayer(index).PokemonBattle.spdef
player(enemyIndex).PokemonInstance(enemySlot).bspd = TempPlayer(index).PokemonBattle.spd
player(enemyIndex).PokemonInstance(enemySlot).hp = TempPlayer(index).PokemonBattle.hp




End Sub


Sub UpdatePVPMainPokemon(ByVal index As Long, ByVal enemyIndex As Long, ByVal enemySlot As Long, ByVal mainSlot As Long)
On Error Resume Next
player(index).PokemonInstance(mainSlot).batk = TempPlayer(enemyIndex).PokemonBattle.atk
player(index).PokemonInstance(mainSlot).bdef = TempPlayer(enemyIndex).PokemonBattle.def
player(index).PokemonInstance(mainSlot).bspatk = TempPlayer(enemyIndex).PokemonBattle.spatk
player(index).PokemonInstance(mainSlot).bspdef = TempPlayer(enemyIndex).PokemonBattle.spdef
player(index).PokemonInstance(mainSlot).bspd = TempPlayer(enemyIndex).PokemonBattle.spd
player(index).PokemonInstance(mainSlot).hp = TempPlayer(enemyIndex).PokemonBattle.hp


End Sub

Sub UpdatePVPEnemyPokemonBattle(ByVal index As Long, ByVal enemyIndex As Long, ByVal enemySlot As Long, ByVal mainSlot As Long)
On Error Resume Next
TempPlayer(enemyIndex).PokemonBattle.atk = player(index).PokemonInstance(mainSlot).batk
TempPlayer(enemyIndex).PokemonBattle.def = player(index).PokemonInstance(mainSlot).bdef
TempPlayer(enemyIndex).PokemonBattle.spatk = player(index).PokemonInstance(mainSlot).bspatk
TempPlayer(enemyIndex).PokemonBattle.spdef = player(index).PokemonInstance(mainSlot).bspdef
TempPlayer(enemyIndex).PokemonBattle.spd = player(index).PokemonInstance(mainSlot).bspd
TempPlayer(enemyIndex).PokemonBattle.hp = player(index).PokemonInstance(mainSlot).hp
TempPlayer(enemyIndex).PokemonBattle.MaxHp = player(index).PokemonInstance(mainSlot).MaxHp
TempPlayer(enemyIndex).PokemonBattle.level = player(index).PokemonInstance(mainSlot).level

TempPlayer(enemyIndex).PokemonBattle.PokemonNumber = player(index).PokemonInstance(mainSlot).PokemonNumber
End Sub

Sub UpdatePVPMainPokemonBattle(ByVal index As Long, ByVal enemyIndex As Long, ByVal enemySlot As Long, ByVal mainSlot As Long)
On Error Resume Next
TempPlayer(index).PokemonBattle.atk = player(enemyIndex).PokemonInstance(enemySlot).batk
TempPlayer(index).PokemonBattle.def = player(enemyIndex).PokemonInstance(enemySlot).bdef
TempPlayer(index).PokemonBattle.spatk = player(enemyIndex).PokemonInstance(enemySlot).bspatk
TempPlayer(index).PokemonBattle.spdef = player(enemyIndex).PokemonInstance(enemySlot).bspdef
TempPlayer(index).PokemonBattle.spd = player(enemyIndex).PokemonInstance(enemySlot).bspd
TempPlayer(index).PokemonBattle.hp = player(enemyIndex).PokemonInstance(enemySlot).hp
TempPlayer(index).PokemonBattle.MaxHp = player(enemyIndex).PokemonInstance(enemySlot).MaxHp
TempPlayer(index).PokemonBattle.level = player(enemyIndex).PokemonInstance(enemySlot).level
TempPlayer(index).PokemonBattle.PokemonNumber = player(enemyIndex).PokemonInstance(enemySlot).PokemonNumber
End Sub


Public Function CanAddRankPoint(ByVal index As Long, ByVal enemyIndex As Long) As Boolean
On Error Resume Next
If GetPlayerDivision(index, GetRankedPoints(index)) >= DIVISION_BRONZE_3 And GetPlayerDivision(index, GetRankedPoints(index)) <= DIVISION_BRONZE_1 Then
If GetPlayerDivision(enemyIndex, GetRankedPoints(enemyIndex)) > DIVISION_GOLD_1 Then
CanAddRankPoint = False
Else
CanAddRankPoint = True
End If
End If

If GetPlayerDivision(index, GetRankedPoints(index)) >= DIVISION_SILVER_3 And GetPlayerDivision(index, GetRankedPoints(index)) <= DIVISION_SILVER_1 Then
If GetPlayerDivision(enemyIndex, GetRankedPoints(enemyIndex)) > DIVISION_PLATINUM_1 Then
CanAddRankPoint = False
Else
CanAddRankPoint = True
End If
End If

If GetPlayerDivision(index, GetRankedPoints(index)) >= DIVISION_GOLD_3 And GetPlayerDivision(index, GetRankedPoints(index)) <= DIVISION_GOLD_1 Then
CanAddRankPoint = True
End If

If GetPlayerDivision(index, GetRankedPoints(index)) >= DIVISION_PLATINUM_3 And GetPlayerDivision(index, GetRankedPoints(index)) <= DIVISION_PLATINUM_1 Then
If GetPlayerDivision(enemyIndex, GetRankedPoints(enemyIndex)) < DIVISION_SILVER_3 Then
CanAddRankPoint = False
Else
CanAddRankPoint = True
End If
End If

If GetPlayerDivision(index, GetRankedPoints(index)) >= DIVISION_DIAMOND_3 And GetPlayerDivision(index, GetRankedPoints(index)) <= DIVISION_DIAMOND_1 Then
If GetPlayerDivision(enemyIndex, GetRankedPoints(enemyIndex)) < DIVISION_GOLD_3 Then
CanAddRankPoint = False
Else
CanAddRankPoint = True
End If
End If


End Function
