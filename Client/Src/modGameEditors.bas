Attribute VB_Name = "modGameEditors"
Option Explicit
' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
On Error Resume Next
Dim i As Long
    InMapEditor = True
    StopPlay
    frmEditor_Map.Visible = True
    Call InitDDSurf("misc", DDSD_Misc, DDS_Misc)
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & map.tileset
    frmEditor_Map.scrlTileSet = map.tileset
    Call EditorMap_BltTileset
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).Name
    Next
    frmEditor_Map.cmbShop.ListIndex = 0
End Sub

Public Sub MapEditorSetTile(ByVal X As Integer, ByVal Y As Integer, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False)
Dim x2 As Integer, y2 As Integer

    If Not multitile Then ' single
        With map.Tile(X, Y)
            ' set layer
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).tileset = frmEditor_Map.scrlTileSet.Value
        End With
    Else ' multitile
        y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            x2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= map.MaxX Then
                    If Y >= 0 And Y <= map.MaxY Then
                        With map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + x2
                            .Layer(CurLayer).Y = EditorTileY + y2
                            .Layer(CurLayer).tileset = frmEditor_Map.scrlTileSet.Value
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
End Sub

Public Sub MapEditorMouseDown(Button As Integer)
Dim i As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer
            Else ' multi tile!
                MapEditorSetTile CurX, CurY, CurLayer, True
            End If
        Else
            With map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .data1 = EditorWarpMap
                    .data2 = EditorWarpX
                    .data3 = EditorWarpY
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .data1 = ItemEditorNum
                    .data2 = ItemEditorValue
                    .data3 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .data1 = 0
                    .data2 = 0
                    .data3 = 0
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .data1 = KeyEditorNum
                    .data2 = KeyEditorTake
                    .data3 = 0
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .data1 = KeyOpenEditorX
                    .data2 = KeyOpenEditorY
                    .data3 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .data1 = ResourceEditorNum
                    .data2 = 0
                    .data3 = 0
                End If
                ' door
                If frmEditor_Map.optDoor.Value Then
                    .Type = TILE_TYPE_DOOR
                    .data1 = EditorWarpMap
                    .data2 = EditorWarpX
                    .data3 = EditorWarpY
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .data1 = SpawnNpcNum
                    .data2 = SpawnNpcDir
                    .data3 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .data1 = EditorShop
                    .data2 = 0
                    .data3 = 0
                End If
                ' battle
                If frmEditor_Map.optBattle.Value Then
                    .Type = TILE_TYPE_BATTLE
                    .data1 = 0
                    .data2 = 0
                    .data3 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .data1 = 0
                    .data2 = 0
                    .data3 = 0
                End If
                'Spawn
                If frmEditor_Map.optSpawn.Value Then
                .Type = TILE_TYPE_SPAWN
                .data1 = 0
                .data2 = 0
                .data3 = 0
                End If
                If frmEditor_Map.optStorage.Value Then
                .Type = TILE_TYPE_STORAGE
                .data1 = 0
                .data2 = 0
                .data3 = 0
                End If
                If frmEditor_Map.optBank.Value Then
                .Type = TILE_TYPE_BANK
                .data1 = 0
                .data2 = 0
                .data3 = 0
                End If
                If frmEditor_Map.optgymblock.Value Then
                .Type = TILE_TYPE_GYMBLOCK
                .data1 = EditorGymBlockNum
                .data2 = EditorGymBlockDir
                .data3 = 0
                End If
                 If frmEditor_Map.optCS.Value Then
                .Type = TILE_TYPE_CUSTOMSCRIPT
                .data1 = EditorGymBlockNum
                .data2 = 0
                .data3 = 0
                End If
            End With
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).tileset = 0
            End With
        Else
            With map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .data1 = 0
                .data2 = 0
                .data3 = 0
            End With

        End If
    End If

    CacheResources
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)


    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
        
        frmEditor_Map.shpSelected.Top = EditorTileY * PIC_Y
        frmEditor_Map.shpSelected.Left = EditorTileX * PIC_X
        
        frmEditor_Map.shpSelected.width = PIC_X
        frmEditor_Map.shpSelected.height = PIC_Y
    End If
    
    ' Error handler

End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
' If

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > frmEditor_Map.picBackSelect.width / PIC_X Then X = frmEditor_Map.picBackSelect.width / PIC_X
        If Y < 0 Then Y = 0
        If Y > frmEditor_Map.picBackSelect.height / PIC_Y Then Y = frmEditor_Map.picBackSelect.height / PIC_Y
        ' find out what to set the width + height of map editor to
        If X > EditorTileX Then ' drag right
            EditorTileWidth = X - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If Y > EditorTileY Then ' drag down
            EditorTileHeight = Y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        frmEditor_Map.shpSelected.width = EditorTileWidth * PIC_X
        frmEditor_Map.shpSelected.height = EditorTileHeight * PIC_Y
    End If

   

End Sub

Public Sub MapEditorTileScroll()
    frmEditor_Map.picBackSelect.Top = (frmEditor_Map.scrlPictureY.Value * PIC_Y) * -1
    frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.Value * PIC_X) * -1
End Sub

Public Sub MapEditorSend()
    Call SendMap
    InMapEditor = False
    PlayMapMusic MapMusic
    frmEditor_Map.Visible = False
    Set DDS_Misc = Nothing
End Sub

Public Sub MapEditorCancel()
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong TCP_CODE
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    frmEditor_Map.Visible = False
    Set DDS_Misc = Nothing
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, GAME_NAME) = vbYes Then
        For X = 0 To map.MaxX
            For Y = 0 To map.MaxY
                map.Tile(X, Y).Layer(CurLayer).X = 0
                map.Tile(X, Y).Layer(CurLayer).Y = 0
                map.Tile(X, Y).Layer(CurLayer).tileset = 0
            Next
        Next
    End If
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, GAME_NAME) = vbYes Then
        For X = 0 To map.MaxX
            For Y = 0 To map.MaxY
                map.Tile(X, Y).Layer(CurLayer).X = EditorTileX
                map.Tile(X, Y).Layer(CurLayer).Y = EditorTileY
                map.Tile(X, Y).Layer(CurLayer).tileset = frmEditor_Map.scrlTileSet.Value
            Next
        Next
    End If
End Sub

Public Sub MapEditorClearAttribs()
    Dim X As Long
    Dim Y As Long

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME) = vbYes Then

        For X = 0 To map.MaxX
            For Y = 0 To map.MaxY
                map.Tile(X, Y).Type = 0
            Next
        Next

    End If

End Sub

Public Sub MapEditorLeaveMap()

    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
    Dim i As Long
    
    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1

    With Item(EditorIndex)
        frmEditor_Item.txtName.text = Trim$(.Name)
        If .pic > frmEditor_Item.scrlPic.Max Then .pic = 0
        frmEditor_Item.scrlPic.Value = .pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .data2
            frmEditor_Item.cmbTool.ListIndex = .data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.Value = .speed
            
            ' loop for stats
            For i = 1 To Stats.stat_count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
        Else
            If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_MASK) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_OUTFIT) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .data2
            frmEditor_Item.cmbTool.ListIndex = .data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.Value = .speed
            
            ' loop for stats
            For i = 1 To Stats.stat_count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
        Else
            frmEditor_Item.fraEquipment.Visible = False
        End If
        End If
        
        
        
        

        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlVitalMod.Value = .data1
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_POKEPOTION) Then
            frmEditor_Item.fraPotionData.Visible = True
            frmEditor_Item.scrlPotionHP.Value = .AddHP
            frmEditor_Item.lblPotionHP.Caption = "+HP : " & .AddHP
        Else
            frmEditor_Item.fraPotionData.Visible = False
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        frmEditor_Item.scrlCatchRate.Value = .CatchRate
        ' loop for stats
        For i = 1 To Stats.stat_count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "None"

        For i = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(i).Name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.Value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Call EditorItem_BltItem
    Call EditorItem_BltPaperdoll
    Item_Changed(EditorIndex) = True
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
End Sub

Public Sub ItemEditorCancel()
    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
End Sub

Public Sub ClearChanged_Item()
    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
    Dim i As Long
    
    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.Name)
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
            frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Call EditorAnim_BltAnim
    Animation_Changed(EditorIndex) = True
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
End Sub

Public Sub AnimationEditorCancel()
    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
End Sub

Public Sub ClearChanged_Animation()
    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
    Dim i As Long
    
    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    With frmEditor_NPC
        .txtName.text = Trim$(NPC(EditorIndex).Name)
        .txtAttackSay.text = Trim$(NPC(EditorIndex).AttackSay)
        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.Max Then NPC(EditorIndex).Sprite = 0
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        .txtSpawnSecs.text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .cmbFaction.ListIndex = NPC(EditorIndex).faction
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtChance.text = CStr(NPC(EditorIndex).DropChance)
        .scrlNum.Value = NPC(EditorIndex).DropItem
        .scrlValue.Value = NPC(EditorIndex).DropItemValue
        .txtHP.text = NPC(EditorIndex).HP
        .txtEXP.text = NPC(EditorIndex).EXP
        .chkCanMove.Value = NPC(EditorIndex).CanMove
        .txtP1.text = Val(NPC(EditorIndex).Paperdoll1)
        .txtP2.text = Val(NPC(EditorIndex).Paperdoll2)
        .txtP3.text = Val(NPC(EditorIndex).Paperdoll3)
        For i = 1 To Stats.stat_count - 1
            .scrlStat(i).Value = NPC(EditorIndex).Stat(i)
        Next
    End With
    
    Call EditorNpc_BltSprite
    NPC_Changed(EditorIndex) = True
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
End Sub

Public Sub NpcEditorCancel()
    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
End Sub

Public Sub ClearChanged_NPC()
    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    frmEditor_Resource.scrlExhaustedPic.Max = NumResources
    frmEditor_Resource.scrlNormalPic.Max = NumResources
    frmEditor_Resource.scrlAnimation.Max = MAX_ANIMATIONS
    
    frmEditor_Resource.txtName.text = Trim$(Resource(EditorIndex).Name)
    frmEditor_Resource.txtMessage.text = Trim$(Resource(EditorIndex).SuccessMessage)
    frmEditor_Resource.txtMessage2.text = Trim$(Resource(EditorIndex).EmptyMessage)
    frmEditor_Resource.cmbType.ListIndex = Resource(EditorIndex).ResourceType
    frmEditor_Resource.scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
    frmEditor_Resource.scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
    frmEditor_Resource.scrlReward.Value = Resource(EditorIndex).ItemReward
    frmEditor_Resource.scrlTool.Value = Resource(EditorIndex).ToolRequired
    frmEditor_Resource.scrlHealth.Value = Resource(EditorIndex).Health
    frmEditor_Resource.scrlRespawn.Value = Resource(EditorIndex).RespawnTime
    frmEditor_Resource.scrlAnimation.Value = Resource(EditorIndex).Animation
    
    Call EditorResource_BltSprite
    
    Resource_Changed(EditorIndex) = True
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
End Sub

Public Sub ResourceEditorCancel()
    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
End Sub

Public Sub ClearChanged_Resource()
    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
End Sub

' ////////////////
' // Pokemon Editor //
' ////////////////
Public Sub PokemonEditorInit()

    If frmEditor_Pokemon.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Pokemon.lstIndex.ListIndex + 1
    frmEditor_Pokemon.Image1.Picture = LoadPicture(App.Path & "\Data Files\graphics\pokemonsprites\" & EditorIndex & ".gif")
    frmEditor_Pokemon.Image2.Picture = LoadPicture(App.Path & "\Data Files\graphics\pokemonsprites\Shiny\" & EditorIndex & ".gif")
    With Pokemon(EditorIndex)
        frmEditor_Pokemon.txtName.text = Trim$(.Name)
        frmEditor_Pokemon.scrlHP.Value = .MaxHp
        frmEditor_Pokemon.scrlPP.Value = .MaxPP
        frmEditor_Pokemon.cmbType.ListIndex = .Type
        frmEditor_Pokemon.cmbType2.ListIndex = .Type2
        If .EvolvesTo > 1 Then
        frmEditor_Pokemon.scrlEvolvePoke.Value = .EvolvesTo
        frmEditor_Pokemon.scrlEvolveLvl.Value = .Evolution
        Else
        frmEditor_Pokemon.scrlEvolvePoke.Value = frmEditor_Pokemon.scrlEvolvePoke.min
        frmEditor_Pokemon.scrlEvolveLvl.Value = 0
        End If
        
        frmEditor_Pokemon.scrlAtk.Value = .ATK
        frmEditor_Pokemon.scrlDef.Value = .DEF
        frmEditor_Pokemon.scrlSpd.Value = .SPD
        frmEditor_Pokemon.scrlSpAtk.Value = .SPATK
        frmEditor_Pokemon.scrlSpDef.Value = .SPDEF
        frmEditor_Pokemon.scrlRareness.Value = .Rareness
        frmEditor_Pokemon.scrlFemalePerc.Value = .PercentFemale
        frmEditor_Pokemon.scrlHappiness.Value = .Happiness
        frmEditor_Pokemon.scrlCatchRate.Value = .CatchRate
        frmEditor_Pokemon.scrlExp.Value = .BaseEXP
        frmEditor_Pokemon.scrlLevel.Value = 1
        frmEditor_Pokemon.scrlMove.Value = .moves(1)
        frmEditor_Pokemon.txtStone.text = Trim$(.Stone)
    End With
    
    Pokemon_Changed(EditorIndex) = True
End Sub

Public Sub PokemonEditorOk()
Dim i As Long

    For i = 1 To MAX_POKEMONS
        If Pokemon_Changed(i) Then
            Call SendSavePokemon(i)
        End If
    Next
    
    Unload frmEditor_Pokemon
    Editor = 0
    ClearChanged_Pokemon
End Sub

Public Sub PokemonEditorCancel()
    Editor = 0
    Unload frmEditor_Pokemon
    ClearChanged_Pokemon
    ClearPokemons
    SendRequestPokemon
End Sub

Public Sub ClearChanged_Pokemon()
    ZeroMemory Pokemon_Changed(1), MAX_POKEMONS * 2 ' 2 = boolean length
End Sub
'Moves Editor
Public Sub MovesEditorInit()
If frmEditor_Moves.Visible = False Then Exit Sub
EditorIndex = frmEditor_Moves.lstIndex.ListIndex + 1
With PokemonMove(EditorIndex)
frmEditor_Moves.txtMoveName.text = Trim$(.Name)
'frmEditor_Moves.cmbType1.ListIndex = .Type
frmEditor_Moves.scrlPower.Value = .power
frmEditor_Moves.scrlPP.Value = .pp
frmEditor_Moves.scrlAccuracy.Value = .accuracy
frmEditor_Moves.HScroll1.Value = .effect
Call ReadText(App.Path & "\Data Files\database\Moves\" & .effect & ".txt", frmEditor_Moves.txtEffectdescription)
frmEditor_Moves.txtDescription.text = .Description
'Category
Select Case Trim$(.Category)
Case "Physical Damage"
'Physical Damage
frmEditor_Moves.cmbCategory.ListIndex = 0
Case "Special Damage"
'Special Damage
frmEditor_Moves.cmbCategory.ListIndex = 1
Case "Status"
'Status
frmEditor_Moves.cmbCategory.ListIndex = 2
End Select

End With
Move_changed(EditorIndex) = True
End Sub
Public Sub MovesEditorOk()
Dim i As Long

    For i = 1 To MAX_MOVES
        If Move_changed(i) Then
            Call SendSaveMove(i)
        End If
    Next
    
    Unload frmEditor_Moves
    Editor = 0
    ClearChanged_Moves
End Sub

Public Sub MovesEditorCancel()
    Editor = 0
    Unload frmEditor_Moves
    ClearChanged_Moves
    ClearMoves
    SendRequestMove
End Sub

Public Sub ClearChanged_Moves()
    ZeroMemory Move_changed(1), 500 * 2
End Sub
' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
    Dim i As Long
    
    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.text = Trim$(Shop(EditorIndex).Name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).Name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
End Sub

Public Sub UpdateShopTrade()
    Dim i As Long
    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                frmEditor_Shop.lstTradeItem.AddItem i & ": " & .itemvalue & "x " & Trim$(Item(.Item).Name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).Name)
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
End Sub

Public Sub ShopEditorCancel()
    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
End Sub

Public Sub ClearChanged_Shop()
    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
    Dim i As Long
    
    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "None"
        For i = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(i).Name)
        Next
        .cmbClass.ListIndex = 0
        
        ' set values
        .txtName.text = Trim$(Spell(EditorIndex).Name)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).dir
        .scrlVital.Value = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
    End With
    
    EditorSpell_BltIcon
    
    Spell_Changed(EditorIndex) = True
End Sub

Public Sub SpellEditorOk()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
End Sub

Public Sub SpellEditorCancel()
    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
End Sub

Public Sub ClearChanged_Spell()
    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
End Sub

Public Sub ClearAttributeDialogue()
frmEditor_Map.fraNpcSpawn.Visible = False
frmEditor_Map.fraResource.Visible = False
frmEditor_Map.fraMapItem.Visible = False
frmEditor_Map.fraMapKey.Visible = False
frmEditor_Map.fraKeyOpen.Visible = False
frmEditor_Map.fraMapWarp.Visible = False
frmEditor_Map.fraShop.Visible = False
End Sub
'//////////STORAGE/////////////By Golf

Public Sub initStorage()
Dim i As Long
frmStorage.lstStorage.Clear
For i = 1 To 250
If StorageInstance(i).PokemonNumber > 0 Then
frmStorage.lstStorage.AddItem (i & ": " & Pokemon(StorageInstance(i).PokemonNumber).Name & " Lvl:" & StorageInstance(i).Level)
Else
frmStorage.lstStorage.AddItem (i & ": Empty")
End If

Next
   
End Sub


