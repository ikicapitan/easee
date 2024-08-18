Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim I As Long
    FindOpenPlayerSlot = 0

    For I = 1 To MAX_PLAYERS

        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If

    Next

End Function

Sub CheckHighlight(ByVal index As Long, ByVal invNum As Long)
Dim reSet As Boolean, Sel1 As Boolean, Sel2 As Boolean
Dim Sel_Index As Long, itemnum As Long
Dim I As Long, II As Long
Dim aiiSelected As Boolean
Dim var1 As String

    ' if selected, unselect
    If Player(index).Inv(invNum).Selected = 1 Then
        Player(index).Inv(invNum).Selected = 0
    Else
        ' highlight the item
        Player(index).Inv(invNum).Selected = 1
        reSet = False
        
        ' see if another one is selected, if so, get ready to use item combo system
        For I = 1 To MAX_INV
            If Player(index).Inv(I).Selected = 1 And I <> invNum Then
            
                ' Run through combos to see if we have one for these items
                For II = 1 To MAX_COMBOS
                    If Combo(II).Item_1 > 1 Or Combo(II).Item_2 > 1 Then
                        itemnum = GetPlayerInvItemNum(index, invNum)
                        
                        ' Check if we have item 1 in the slot we just clicked
                        If itemnum = Combo(II).Item_1 Then Sel1 = True
                        
                        ' Check if we have item 2 in the slot we just clicked
                        If itemnum = Combo(II).Item_2 Then Sel1 = True
                        
                        
                        itemnum = GetPlayerInvItemNum(index, I)
                            
                        ' Check if we have item 1 in the other slot
                        If itemnum = Combo(II).Item_1 Then Sel2 = True
                            
                        ' Check if we have item 2 in the other slot
                        If itemnum = Combo(II).Item_2 Then Sel2 = True
                    End If
                    
                    
                    ' Leave the loop if we found a combo
                    If Sel1 = True And Sel2 = True Then
                        Sel_Index = II
                        Exit For
                    End If
                    
                    Sel1 = False
                    Sel2 = False
                Next II
                
                ' If both items are part of a combo then we're moving on
                If Sel1 = True And Sel2 = True Then
                    aiiSelected = True
                End If
                
                Player(index).Inv(I).Selected = 0
                reSet = True
            End If
        Next I
    End If
    
    'use item combo system
    If aiiSelected Then
        ' Check requirements
        If Combo(Sel_Index).Level > 0 And GetPlayerLevel(index) < Combo(Sel_Index).Level Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M1")
            Call PlayerMsg(index, var1 & Combo(Sel_Index).Level & ".", BrightRed) ' MateoD
            GoTo Continue   '[TRADESKILL]M1
        End If
        If Combo(Sel_Index).Skill > 0 And GetPlayerSkillLevel(index, Combo(Sel_Index).Skill) < Combo(Sel_Index).SkillLevel Then
            Dim var2
            
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M2")
var2 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M3")
            Call PlayerMsg(index, var1 & Trim$(Skill(Combo(Sel_Index).Skill).Name) & var2 & Combo(Sel_Index).SkillLevel & ".", BrightRed) ' MateoD
            GoTo Continue       '[TRADESKILL]M2                                      [TRADESKILL]M3
        End If
        If Combo(Sel_Index).ReqItem1 > 0 And HasItems(index, Combo(Sel_Index).ReqItem1, Combo(Sel_Index).ReqItemVal1) = False Then
            If Item(Combo(Sel_Index).ReqItem1).Type = ITEM_TYPE_CURRENCY Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M4")
                Call PlayerMsg(index, var1 & Combo(Sel_Index).ReqItemVal1 & " " & Trim$(Item(Combo(Sel_Index).ReqItem1).Name) & ".", BrightRed) ' MateoD
            Else                        '[TRADESKILL]M4
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M4")
                Call PlayerMsg(index, var1 & CheckGrammar(Trim$(Item(Combo(Sel_Index).ReqItem1).Name)) & ".", BrightRed) ' MateoD
            End If                        '[TRADESKILL]M4
            GoTo Continue
        End If
        If Combo(Sel_Index).ReqItem2 > 0 And HasItems(index, Combo(Sel_Index).ReqItem2, Combo(Sel_Index).ReqItemVal2) = False Then
            If Item(Combo(Sel_Index).ReqItem2).Type = ITEM_TYPE_CURRENCY Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M4")
                Call PlayerMsg(index, var1 & Combo(Sel_Index).ReqItemVal2 & " " & Trim$(Item(Combo(Sel_Index).ReqItem2).Name) & ".", BrightRed) ' MateoD
            Else                        '[TRADESKILL]M4
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TRADESKILL", "M4")
                Call PlayerMsg(index, var1 & CheckGrammar(Trim$(Item(Combo(Sel_Index).ReqItem2).Name)) & ".", BrightRed) ' MateoD
            End If                        '[TRADESKILL]M4
            GoTo Continue
        End If
                
    
        ' Take items
        If Combo(Sel_Index).Take_Item1 = 1 Then Call TakeInvItem(index, Combo(Sel_Index).Item_1, 1)
        If Combo(Sel_Index).Take_Item2 = 1 Then Call TakeInvItem(index, Combo(Sel_Index).Item_2, 1)
        If Combo(Sel_Index).Take_ReqItem1 = 1 Then Call TakeInvItem(index, Combo(Sel_Index).ReqItem1, 1)
        If Combo(Sel_Index).Take_ReqItem2 = 1 Then Call TakeInvItem(index, Combo(Sel_Index).ReqItem2, 1)
                
        ' Give items
        For I = 1 To MAX_COMBO_GIVEN
            If Combo(Sel_Index).Item_Given(I) > 0 Then
                If Item(Combo(Sel_Index).Item_Given(I)).Type = ITEM_TYPE_CURRENCY Then
                    Call GiveInvItem(index, Combo(Sel_Index).Item_Given(I), Combo(Sel_Index).Item_Given_Val(I))
                Else
                    For II = 1 To Combo(Sel_Index).Item_Given_Val(I)
                        Call GiveInvItem(index, Combo(Sel_Index).Item_Given(I), 1)
                    Next II
                End If
            End If
        Next I
        
        If Combo(Sel_Index).GiveSkill > 0 Then
            Call SetPlayerSkillExp(index, Combo(Sel_Index).GiveSkill, Combo(Sel_Index).GiveSkill_Exp)
            Call SendPlayerData(index)
            Call PlayerMsg(index, "Ganas " & Combo(Sel_Index).GiveSkill_Exp & " " & Trim$(Skill(Combo(Sel_Index).GiveSkill).Name) & " experiencia.", Cyan)
        End If
    End If
Continue:
    
    If reSet Then
        ' Remove all highlights
        For I = 1 To MAX_INV
            Player(index).Inv(I).Selected = 0
            SendHighlight index, I
        Next I
    End If
End Sub

Sub RemoveFriend(ByVal index As Long, ByVal fName As String)
Dim I As Long, Place As Long, pI As Long, fOther As String
    pI = FindPlayer(fName)
    fOther = GetPlayerName(index)
    
    ' Do the first player
    For I = 1 To GetPlayerFriends(index)
        If GetPlayerFriendName(index, I) = fName Then
            Place = I
            Call SetPlayerFriends(index, -1)
        End If
        
        If Place > 0 And Place < I Then
            Call SetPlayerFriendName(index, I - 1, GetPlayerFriendName(index, I))
            SetPlayerFriendName index, I
        End If
    Next I
    Place = 0
    
    ' Do the other player
    For I = 1 To GetPlayerFriends(pI)
        If GetPlayerFriendName(pI, I) = fOther Then
            Place = I
            SetPlayerFriendName pI, I
            Call SetPlayerFriends(pI, -1)
        End If
        
        If Place > 0 And Place < I Then
            GetPlayerFriendName(pI, I - 1) = GetPlayerFriendName(pI, I)
        End If
    Next I
End Sub

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim I As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For I = 1 To MAX_MAP_ITEMS

        If MapItem(mapnum, I).num = 0 Then
            FindOpenMapItemSlot = I
            Exit Function
        End If

    Next

End Function

Function TotalOnlinePlayers() As Long
    Dim I As Long
    TotalOnlinePlayers = 0

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If

    Next

End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(I), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString)
    Dim I As Long

    ' Check for subscript out of range
    If itemnum < 1 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    I = FindOpenMapItemSlot(mapnum)
    Call SpawnItemSlot(I, itemnum, ItemVal, mapnum, x, y, playerName)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal itemnum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or itemnum < 0 Or itemnum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I <> 0 Then
        If itemnum >= 0 And itemnum <= MAX_ITEMS Then
            MapItem(mapnum, I).playerName = playerName
            MapItem(mapnum, I).playerTimer = GetTickCount + ITEM_SPAWN_TIME
            MapItem(mapnum, I).canDespawn = canDespawn
            MapItem(mapnum, I).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(mapnum, I).num = itemnum
            MapItem(mapnum, I).Value = ItemVal
            MapItem(mapnum, I).x = x
            MapItem(mapnum, I).y = y
            ' send to map
            SendSpawnItemToMap mapnum, I
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next

End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapnum).Tile(x, y).Data1).Stackable > 0 And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If

        Next
    Next

End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Dim temp As Long
    
    'Make sure (High) is actually the high number
    If Low > High Then
        temp = High
        High = Low
        Low = temp
    End If
    
    'continue
    Random = ((High - Low + 1) * Rnd) + Low
End Function

Public Sub SpawnNpc(ByVal mapNpcNum As Long, ByVal mapnum As Long, Optional ForcedSpawn As Boolean = False)
On Error GoTo error:
    Dim buffer As clsBuffer
    Dim npcNum As Long
    Dim I As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean
    Dim HPRndNum As Long
    Dim NText As String
    Dim SEP_CHAR As String * 1

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Sub
    npcNum = Map(mapnum).NPC(mapNpcNum)
    If ForcedSpawn = False And Map(mapnum).NpcSpawnType(mapNpcNum) = 1 Then npcNum = 0
    If npcNum > 0 Then
        NText = Replace$(NPC(npcNum).Name, SEP_CHAR, vbNullString)
        If Len(NText) < 1 Then Exit Sub
    
        MapNpc(mapnum).NPC(mapNpcNum).num = npcNum
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0 ' clear
        
        If NPC(mapNpcNum).RandHP = 0 Then
            MapNpc(mapnum).NPC(mapNpcNum).HPSetTo = GetNpcMaxVital(npcNum, Vitals.HP)
            MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).HPSetTo
        Else
            HPRndNum = rand(NPC(npcNum).HPMin, NPC(npcNum).HP)
            MapNpc(mapnum).NPC(mapNpcNum).HPSetTo = HPRndNum
            MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = HPRndNum
        End If
        
        MapNpc(mapnum).NPC(mapNpcNum).Dir = Int(Rnd * 4)
        
        'Check if theres a spawn tile for the specific npc
        For x = 0 To Map(mapnum).MaxX
            For y = 0 To Map(mapnum).MaxY
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(mapnum).Tile(x, y).Data1 = mapNpcNum Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
                        MapNpc(mapnum).NPC(mapNpcNum).Dir = Map(mapnum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For I = 1 To 100
                x = Random(0, Map(mapnum).MaxX)
                y = Random(0, Map(mapnum).MaxY)
    
                If x > Map(mapnum).MaxX Then x = Map(mapnum).MaxX
                If y > Map(mapnum).MaxY Then y = Map(mapnum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(mapnum, x, y) Then
                    MapNpc(mapnum).NPC(mapNpcNum).x = x
                    MapNpc(mapnum).NPC(mapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(mapnum).MaxX
                For y = 0 To Map(mapnum).MaxY

                    If NpcTileIsOpen(mapnum, x, y) Then
                        MapNpc(mapnum).NPC(mapNpcNum).x = x
                        MapNpc(mapnum).NPC(mapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnNpc
            buffer.WriteLong mapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).num
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).HPSetTo
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
            UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, True
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    Else
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0 ' clear
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong mapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
    End If
error:
End Sub

Public Sub SpawnMapEventsFor(index As Long, mapnum As Long)
Dim I As Long, x As Long, y As Long, z As Long, spawncurrentevent As Boolean, P As Long
Dim buffer As clsBuffer
    
    TempPlayer(index).EventMap.CurrentEvents = 0
    ReDim TempPlayer(index).EventMap.EventPages(0)
    
    If Map(mapnum).EventCount <= 0 Then Exit Sub
    For I = 1 To Map(mapnum).EventCount
        If Map(mapnum).Events(I).PageCount > 0 Then
            For z = Map(mapnum).Events(I).PageCount To 1 Step -1
                With Map(mapnum).Events(I).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Player(index).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Player(index).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(index, .HasItemIndex) < .HasItemIndex Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(mapnum).Events(I).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        'spawn the event... send data to player
                        TempPlayer(index).EventMap.CurrentEvents = TempPlayer(index).EventMap.CurrentEvents + 1
                        ReDim Preserve TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                        With TempPlayer(index).EventMap.EventPages(TempPlayer(index).EventMap.CurrentEvents)
                            If Map(mapnum).Events(I).Pages(z).GraphicType = 1 Then
                                Select Case Map(mapnum).Events(I).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            .GraphicNum = Map(mapnum).Events(I).Pages(z).Graphic
                            .GraphicType = Map(mapnum).Events(I).Pages(z).GraphicType
                            .GraphicX = Map(mapnum).Events(I).Pages(z).GraphicX
                            .GraphicY = Map(mapnum).Events(I).Pages(z).GraphicY
                            .GraphicX2 = Map(mapnum).Events(I).Pages(z).GraphicX2
                            .GraphicY2 = Map(mapnum).Events(I).Pages(z).GraphicY2
                            Select Case Map(mapnum).Events(I).Pages(z).MoveSpeed
                                Case 0
                                    .movementspeed = 2
                                Case 1
                                    .movementspeed = 3
                                Case 2
                                    .movementspeed = 4
                                Case 3
                                    .movementspeed = 6
                                Case 4
                                    .movementspeed = 12
                                Case 5
                                    .movementspeed = 24
                            End Select
                            If Map(mapnum).Events(I).Global Then
                                .x = TempEventMap(mapnum).Events(I).x
                                .y = TempEventMap(mapnum).Events(I).y
                                .Dir = TempEventMap(mapnum).Events(I).Dir
                                .MoveRouteStep = TempEventMap(mapnum).Events(I).MoveRouteStep
                            Else
                                .x = Map(mapnum).Events(I).x
                                .y = Map(mapnum).Events(I).y
                                .MoveRouteStep = 0
                            End If
                            .Position = Map(mapnum).Events(I).Pages(z).Position
                            .eventID = I
                            .pageID = z
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(mapnum).Events(I).Pages(z).MoveType
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(mapnum).Events(I).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(mapnum).Events(I).Pages(z).MoveRouteCount)
                                If Map(mapnum).Events(I).Pages(z).MoveRouteCount > 0 Then
                                    For P = 0 To Map(mapnum).Events(I).Pages(z).MoveRouteCount
                                        .MoveRoute(P) = Map(mapnum).Events(I).Pages(z).MoveRoute(P)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(mapnum).Events(I).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(mapnum).Events(I).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(mapnum).Events(I).Pages(z).MoveFreq
                            .MoveSpeed = Map(mapnum).Events(I).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(mapnum).Events(I).Pages(z).WalkAnim
                            .WalkThrough = Map(mapnum).Events(I).Pages(z).WalkThrough
                            .ShowName = Map(mapnum).Events(I).Pages(z).ShowName
                            .FixedDir = Map(mapnum).Events(I).Pages(z).DirFix
                            
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For I = 1 To TempPlayer(index).EventMap.CurrentEvents
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnEvent
            buffer.WriteLong I
            With TempPlayer(index).EventMap.EventPages(I)
                buffer.WriteString Map(GetPlayerMap(index)).Events(I).Name
                buffer.WriteLong .Dir
                buffer.WriteLong .GraphicNum
                buffer.WriteLong .GraphicType
                buffer.WriteLong .GraphicX
                buffer.WriteLong .GraphicX2
                buffer.WriteLong .GraphicY
                buffer.WriteLong .GraphicY2
                buffer.WriteLong .movementspeed
                buffer.WriteLong .x
                buffer.WriteLong .y
                buffer.WriteLong .Position
                buffer.WriteLong .Visible
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkAnim
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).DirFix
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).WalkThrough
                buffer.WriteLong Map(mapnum).Events(.eventID).Pages(.pageID).ShowName
            End With
            SendDataTo index, buffer.ToArray
            Set buffer = Nothing
        Next
    End If
End Sub

Public Function NpcTileIsOpen(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long) As Boolean
On Error GoTo error:
    Dim LoopI As Long
    NpcTileIsOpen = True

    If PlayersOnMap(mapnum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = mapnum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(mapnum).NPC(LoopI).num > 0 Then
            If MapNpc(mapnum).NPC(LoopI).x = x Then
                If MapNpc(mapnum).NPC(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next
    
    For LoopI = 1 To TempEventMap(mapnum).EventCount
        If TempEventMap(mapnum).Events(LoopI).active = 1 Then
            If MapNpc(mapnum).NPC(LoopI).x = TempEventMap(mapnum).Events(LoopI).x Then
                If MapNpc(mapnum).NPC(LoopI).y = TempEventMap(mapnum).Events(LoopI).y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(mapnum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If
error:
End Function

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, mapnum)
    Next
    
    CacheMapBlocks mapnum

End Sub

Sub SpawnAllMapNpcs()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapNpcs(I)
    Next

End Sub

Sub SpawnAllMapGlobalEvents()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnGlobalEvents(I)
    Next

End Sub

Sub SpawnGlobalEvents(ByVal mapnum As Long)
    Dim I As Long, z As Long
    
    If Map(mapnum).EventCount > 0 Then
        TempEventMap(mapnum).EventCount = 0
        ReDim TempEventMap(mapnum).Events(0)
        For I = 1 To Map(mapnum).EventCount
            TempEventMap(mapnum).EventCount = TempEventMap(mapnum).EventCount + 1
            ReDim Preserve TempEventMap(mapnum).Events(0 To TempEventMap(mapnum).EventCount)
            If Map(mapnum).Events(I).PageCount > 0 Then
                If Map(mapnum).Events(I).Global = 1 Then
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).x = Map(mapnum).Events(I).x
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).y = Map(mapnum).Events(I).y
                    If Map(mapnum).Events(I).Pages(1).GraphicType = 1 Then
                        Select Case Map(mapnum).Events(I).Pages(1).GraphicY
                            Case 0
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).Dir = DIR_DOWN
                    End If
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).active = 1
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = Map(mapnum).Events(I).Pages(1).MoveType
                    
                    If TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveType = 2 Then
                        TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRouteCount = Map(mapnum).Events(I).Pages(1).MoveRouteCount
                        ReDim TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(0 To Map(mapnum).Events(I).Pages(1).MoveRouteCount)
                        For z = 0 To Map(mapnum).Events(I).Pages(1).MoveRouteCount
                            TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveRoute(z) = Map(mapnum).Events(I).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).RepeatMoveRoute = Map(mapnum).Events(I).Pages(1).RepeatMoveRoute
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).IgnoreIfCannotMove = Map(mapnum).Events(I).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveFreq = Map(mapnum).Events(I).Pages(1).MoveFreq
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).MoveSpeed = Map(mapnum).Events(I).Pages(1).MoveSpeed
                    
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkThrough = Map(mapnum).Events(I).Pages(1).WalkThrough
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).FixedDir = Map(mapnum).Events(I).Pages(1).DirFix
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).WalkingAnim = Map(mapnum).Events(I).Pages(1).WalkAnim
                    TempEventMap(mapnum).Events(TempEventMap(mapnum).EventCount).ShowName = Map(mapnum).Events(I).Pages(1).ShowName
                    
                End If
            End If
        Next
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim I As Long
    Dim N As Long
    Dim x As Long
    Dim y As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    x = MapNpc(mapnum).NPC(mapNpcNum).x
    y = MapNpc(mapnum).NPC(mapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                N = Map(mapnum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(I) = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> mapNpcNum) And (MapNpc(mapnum).NPC(I).num > 0) And (MapNpc(mapnum).NPC(I).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(I).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                N = Map(mapnum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = MapNpc(mapnum).NPC(mapNpcNum).x) And (GetPlayerY(I) = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> mapNpcNum) And (MapNpc(mapnum).NPC(I).num > 0) And (MapNpc(mapnum).NPC(I).x = MapNpc(mapnum).NPC(mapNpcNum).x) And (MapNpc(mapnum).NPC(I).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                N = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (GetPlayerY(I) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> mapNpcNum) And (MapNpc(mapnum).NPC(I).num > 0) And (MapNpc(mapnum).NPC(I).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1) And (MapNpc(mapnum).NPC(I).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                N = Map(mapnum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (GetPlayerY(I) = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (I <> mapNpcNum) And (MapNpc(mapnum).NPC(I).num > 0) And (MapNpc(mapnum).NPC(I).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1) And (MapNpc(mapnum).NPC(I).y = MapNpc(mapnum).NPC(mapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

End Function

Sub NpcMove(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long, ByVal movement As Long)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, False

    Select Case Dir
        Case DIR_UP
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong mapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_DOWN
            MapNpc(mapnum).NPC(mapNpcNum).y = MapNpc(mapnum).NPC(mapNpcNum).y + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong mapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_LEFT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x - 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong mapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
        Case DIR_RIGHT
            MapNpc(mapnum).NPC(mapNpcNum).x = MapNpc(mapnum).NPC(mapNpcNum).x + 1
            Set buffer = New clsBuffer
            buffer.WriteLong SNpcMove
            buffer.WriteLong mapNpcNum
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).x
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).y
            buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Dir
            buffer.WriteLong movement
            SendDataToMap mapnum, buffer.ToArray()
            Set buffer = Nothing
    End Select
    
    UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, True

End Sub

Sub NpcDir(ByVal mapnum As Long, ByVal mapNpcNum As Long, ByVal Dir As Long)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNpc(mapnum).NPC(mapNpcNum).Dir = Dir
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDir
    buffer.WriteLong mapNpcNum
    buffer.WriteLong Dir
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim I As Long
    Dim N As Long
    N = 0

    For I = 1 To Player_HighIndex

        If isPlaying(I) And GetPlayerMap(I) = mapnum Then
            N = N + 1
        End If

    Next

    GetTotalMapPlayers = N
End Function

Sub ClearTempTiles()
    Dim I As Long

    For I = 1 To MAX_MAPS
        ClearTempTile I
    Next

End Sub

Sub ClearTempTile(ByVal mapnum As Long)
    Dim y As Long
    Dim x As Long
    temptile(mapnum).DoorTimer = 0
    ReDim temptile(mapnum).DoorOpen(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            temptile(mapnum).DoorOpen(x, y) = NO
        Next
    Next

End Sub

Public Sub CacheResources(ByVal mapnum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    Resource_Count = 0

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mapnum).ResourceData(0 To Resource_Count)
                ResourceCache(mapnum).ResourceData(Resource_Count).x = x
                ResourceCache(mapnum).ResourceData(Resource_Count).y = y
                ResourceCache(mapnum).ResourceData(Resource_Count).cur_health = Resource(Map(mapnum).Tile(x, y).Data1).health
            End If

        Next
    Next

    ResourceCache(mapnum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
Dim OldNum As Long
Dim OldValue As Long
Dim NewNum As Long
Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)
Dim var1 As String

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error'd
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        If Item(GetPlayerEquipment(index, EqSlot)).Stackable > 0 Then
            GiveInvItem index, GetPlayerEquipment(index, EqSlot), 1
        Else
            GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        End If
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "OTROS", "M1")
        PlayerMsg index, var1 & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound    [OTROS]M1
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendWornEquipment index
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "OTROS", "M2")
        PlayerMsg index, var1, BrightRed
    End If              '[OTROS]M2

End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
      CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
      Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "un " & Word Else CheckGrammar = "un " & Word
    Else
        If Caps Then CheckGrammar = "un " & Word Else CheckGrammar = "un " & Word
    End If
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
Dim nVal As Long
    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True
End Function

Function FindNPCInRange(ByVal index As Long, ByVal mapnum As Long, ByVal Range As Long)
Dim x As Long, y As Long
Dim npcX As Long, npcY As Long
Dim I As Long

    FindNPCInRange = 0
    
    For I = 1 To MAX_MAP_NPCS
        x = GetPlayerX(index)
        y = GetPlayerY(index)
        npcX = MapNpc(mapnum).NPC(I).x
        npcY = MapNpc(mapnum).NPC(I).y
    
        If isInRange(Range, x, y, npcX, npcY) Then
            If MapNpc(mapnum).NPC(I).num > 0 Then
                If NPC(MapNpc(mapnum).NPC(I).num).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(MapNpc(mapnum).NPC(I).num).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
                    FindNPCInRange = I
                    Exit Function
                End If
            End If
        End If
        
    Next I
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
End Function

Public Function rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    rand = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
Dim partyNum As Long, I As Long
Dim var1 As String

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there's more than 2 people
        If Party(partyNum).MemberCount > 2 Then
        
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(I) > 0 And Party(partyNum).Member(I) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(I)
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M1")
                        PartyMsg partyNum, GetPlayerName(I) & var1, BrightBlue ' MateoD
                        Exit For                            '[GRUPO]M1
                    End If
                Next
                ' leave party
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M2")
                PartyMsg partyNum, GetPlayerName(index) & var1, BrightRed ' MateoD
                ' remove from array                         [GRUPO]M2
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(I) = index Then
                        Party(partyNum).Member(I) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                        End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M2")
                PartyMsg partyNum, GetPlayerName(index) & var1, BrightRed ' MateoD
                ' remove from array                         [GRUPO]M2
                For I = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(I) = index Then
                        Party(partyNum).Member(I) = 0
                        TempPlayer(index).inParty = 0
                        TempPlayer(index).partyInvite = 0
                        Exit For
                    End If
                Next
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M3")
            PartyMsg partyNum, var1, BrightRed ' [GRUPO]M3
            ' clear out everyone's party
            For I = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(I)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).partyInvite = 0
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, I As Long
Dim var1 As String

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not isPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they're not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they've already got a request for trade/party
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M4")
        PlayerMsg index, var1, BrightRed  '[GRUPO]M4
        ' exit out early
        Exit Sub
    End If
    ' make syure they're not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they're already in a party
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M5")
        PlayerMsg index, var1, BrightRed ' [GRUPO]M5
        'exit out early
        Exit Sub
    End If
    
    ' check if we're in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we're the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For I = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(I) = 0 Then
                    ' send the invitation
                    SendPartyInvite targetPlayer, index
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M6")
                    PlayerMsg index, var1, Pink ' [GRUPO]M6
                    Exit Sub
                End If
            Next
            ' no room
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M7")
            PlayerMsg index, var1, BrightRed ' [GRUPO]M7
            Exit Sub
        Else
            ' not the leader
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M8")
            PlayerMsg index, var1, BrightRed ' [GRUPO]M8
            Exit Sub
        End If
    Else
        ' not in a party - doesn't matter!
        SendPartyInvite targetPlayer, index
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M6")
        PlayerMsg index, var1, Pink ' [GRUPO]M6
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
Dim partyNum As Long, I As Long
Dim var1 As String

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For I = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(I) = 0 Then
                'add to the party
                Party(partyNum).Member(I) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                SendPartyVitals partyNum, targetPlayer
                ' let everyone know they've joined
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M9")
                PartyMsg partyNum, GetPlayerName(targetPlayer) & var1, Pink ' MateoD
                ' add them in                                       [GRUPO]M9
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M7")
        PlayerMsg index, var1, BrightRed ' [GRUPO]M7
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M7")
        PlayerMsg targetPlayer, var1, BrightRed ' [GRUPO]M7
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For I = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(I).Leader > 0 Then
                partyNum = I
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it's created
        PartyMsg partyNum, "Grupo creado.", BrightGreen
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M9")
        PartyMsg partyNum, GetPlayerName(index) & var1, Pink ' [GRUPO]M9
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M9")
        PartyMsg partyNum, GetPlayerName(targetPlayer) & var1, Pink ' [GRUPO]M9
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
Dim var1 As String

                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M10")
    PlayerMsg index, GetPlayerName(targetPlayer) & var1, BrightRed ' [GRUPO]M10
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GRUPO", "M11")
    PlayerMsg targetPlayer, var1, BrightRed ' [GRUPO]M11
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
Dim I As Long, highIndex As Long, x As Long
    ' find the high index
    For I = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(I) > 0 Then
            highIndex = I
            Exit For
        End If
    Next
    ' count the members
    For I = 1 To MAX_PARTY_MEMBERS
        ' we've got a blank member
        If Party(partyNum).Member(I) = 0 Then
            ' is it lower than the high index?
            If I < highIndex Then
                ' move everyone down a slot
                For x = I To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we've reached the max
        If I = MAX_PARTY_MEMBERS Then
            If highIndex = I Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we're here it means that we need to re-count again
    Party_CountMembers partyNum
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal EXP As Long, ByVal index As Long, ByVal mapnum As Long)
Dim expShare As Long, leftOver As Long, I As Long, tmpIndex As Long, LoseMemberCount As Byte

    ' check if it's worth sharing
    If Not EXP >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, EXP
        Exit Sub
    End If
    
    ' check members in outhers maps
    For I = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(I)
        If tmpIndex > 0 Then
            If IsConnected(tmpIndex) And isPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) <> mapnum Then
                    LoseMemberCount = LoseMemberCount + 1
                End If
            End If
        End If
    Next I
    
    ' find out the equal share
    expShare = EXP \ (Party(partyNum).MemberCount - LoseMemberCount)
    leftOver = EXP Mod (Party(partyNum).MemberCount - LoseMemberCount)
    
    ' loop through and give everyone exp
    For I = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(I)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And isPlaying(tmpIndex) Then
                If GetPlayerMap(tmpIndex) = mapnum Then
                    ' give them their share
                    GivePlayerEXP tmpIndex, expShare
                End If
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(rand(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal EXP As Long)
    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + EXP)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & EXP & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we've leveled
    CheckPlayerLevelUp index
End Sub

Public Sub GivePlayerCombatEXP(ByVal index As Long, ByVal skillType As Byte, ByVal EXP As Long)
    If EXP < 0 Then Exit Sub
    If Player(index).Combat(Item(skillType).CombatTypeReq).Level = MAX_COMBAT_LEVEL Then Exit Sub
    Call SetPlayerCombatExp(index, Item(skillType).CombatTypeReq, GetPlayerCombatExp(index, Item(skillType).CombatTypeReq) + EXP)
    SendCombatEXP index
    CheckCombatLevelUp index, Item(skillType).CombatTypeReq
End Sub

Function CanEventMove(index As Long, ByVal mapnum As Long, x As Long, y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional globalevent As Boolean = False) As Boolean
    Dim I As Long
    Dim N As Long, z As Long

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If
    CanEventMove = True
    
    

    Select Case Dir
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                N = Map(mapnum).Tile(x, y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                
                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN And N <> TILE_TYPE_NPCAVOID Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = x) And (GetPlayerY(I) = y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(I).x = x) And (MapNpc(mapnum).NPC(I).y = y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(mapnum).MaxY Then
                N = Map(mapnum).Tile(x, y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = x) And (GetPlayerY(I) = y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(I).x = x) And (MapNpc(mapnum).NPC(I).y = y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x) And (TempEventMap(mapnum).Events(z).y = y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                N = Map(mapnum).Tile(x - 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = x - 1) And (GetPlayerY(I) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(I).x = x - 1) And (MapNpc(mapnum).NPC(I).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x - 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(mapnum).MaxX Then
                N = Map(mapnum).Tile(x + 1, y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If N <> TILE_TYPE_WALKABLE And N <> TILE_TYPE_ITEM And N <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For I = 1 To Player_HighIndex
                    If isPlaying(I) Then
                        If (GetPlayerMap(I) = mapnum) And (GetPlayerX(I) = x + 1) And (GetPlayerY(I) = y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For I = 1 To MAX_MAP_NPCS
                    If (MapNpc(mapnum).NPC(I).x = x + 1) And (MapNpc(mapnum).NPC(I).y = y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If globalevent = True Then
                    If TempEventMap(mapnum).EventCount > 0 Then
                        For z = 1 To TempEventMap(mapnum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(mapnum).Events(z).x = x + 1) And (TempEventMap(mapnum).Events(z).y = y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
                            If (TempPlayer(index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (TempPlayer(index).EventMap.EventPages(z).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1) And (TempPlayer(index).EventMap.EventPages(z).y = TempPlayer(index).EventMap.EventPages(eventID).y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If isDirBlocked(Map(mapnum).Tile(x, y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

    End Select

End Function

Sub EventDir(playerindex As Long, ByVal mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional globalevent As Boolean = False)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(playerindex).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventDir
    buffer.WriteLong eventID
    If globalevent Then
        buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
    Else
        buffer.WriteLong TempPlayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub EventMove(index As Long, mapnum As Long, ByVal eventID As Long, ByVal Dir As Long, movementspeed As Long, Optional globalevent As Boolean = False)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapnum <= 0 Or mapnum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If
    
    If globalevent Then
        If Map(mapnum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(mapnum).Events(eventID).Dir = Dir
        UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, False
    Else
        If Map(mapnum).Events(eventID).Pages(TempPlayer(index).EventMap.EventPages(eventID).pageID).DirFix = 0 Then TempPlayer(index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
            
        Case DIR_DOWN
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).y = TempEventMap(mapnum).Events(eventID).y + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).y = TempPlayer(index).EventMap.EventPages(eventID).y + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
        Case DIR_LEFT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x - 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
        Case DIR_RIGHT
            If globalevent Then
                TempEventMap(mapnum).Events(eventID).x = TempEventMap(mapnum).Events(eventID).x + 1
                UpdateMapBlock mapnum, TempEventMap(mapnum).Events(eventID).x, TempEventMap(mapnum).Events(eventID).y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).x
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(mapnum).Events(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                TempPlayer(index).EventMap.EventPages(eventID).x = TempPlayer(index).EventMap.EventPages(eventID).x + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).x
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).y
                buffer.WriteLong Dir
                buffer.WriteLong TempPlayer(index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong movementspeed
                If globalevent Then
                    SendDataToMap mapnum, buffer.ToArray()
                Else
                    SendDataTo index, buffer.ToArray
                End If
                Set buffer = Nothing
            End If
    End Select

End Sub



Public Sub HandleProjecTile(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim x As Long, y As Long, I As Long

    ' check for subscript out of range
    If index < 1 Or index > MAX_PLAYERS Or PlayerProjectile < 1 Or PlayerProjectile > MAX_PLAYER_PROJECTILES Then Exit Sub
        
    ' check to see if it's time to move the Projectile
    If GetTickCount > TempPlayer(index).ProjecTile(PlayerProjectile).TravelTime Then
        With TempPlayer(index).ProjecTile(PlayerProjectile)
            ' set next travel time and the current position and then set the actual direction based on RMXP arrow tiles.
            Select Case .Direction
                ' down
                Case DIR_DOWN
                    .y = .y + 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) + .Range) + 1 Then
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                ' up
                Case DIR_UP
                    .y = .y - 1
                    ' check if they reached maxrange
                    If .y = (GetPlayerY(index) - .Range) - 1 Then
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                ' right
                Case DIR_RIGHT
                    .x = .x + 1
                    ' check if they reached max range
                    If .x = (GetPlayerX(index) + .Range) + 1 Then
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                ' left
                Case DIR_LEFT
                    .x = .x - 1
                    ' check if they reached maxrange
                    If .x = (GetPlayerX(index) - .Range) - 1 Then
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
            End Select
            .TravelTime = GetTickCount + .Speed
        End With
    End If
    
    x = TempPlayer(index).ProjecTile(PlayerProjectile).x
    y = TempPlayer(index).ProjecTile(PlayerProjectile).y
    
    ' check if left map
    If x > Map(GetPlayerMap(index)).MaxX Or y > Map(GetPlayerMap(index)).MaxY Or x < 0 Or y < 0 Then
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
    ' check if hit player
    For I = 1 To Player_HighIndex
        ' make sure they're actually playing
        If isPlaying(I) Then
            ' check coordinates
            If x = Player(I).x And y = GetPlayerY(I) Then
                ' make sure it's not the attacker
                If Not x = Player(index).x Or Not y = GetPlayerY(index) Then
                    ' check if player can attack
                    If CanPlayerAttackPlayer(index, I, False, True) = True Then
                        ' attack the player and kill the project tile
                        PlayerAttackPlayer index, I, TempPlayer(index).ProjecTile(PlayerProjectile).Damage
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    Else
                        ClearProjectile index, PlayerProjectile
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' check for npc hit
    For I = 1 To MAX_MAP_NPCS
        If x = MapNpc(GetPlayerMap(index)).NPC(I).x And y = MapNpc(GetPlayerMap(index)).NPC(I).y Then
            ' they're hit, remove it and deal that damage ;)
            If CanPlayerAttackNpc(index, I, True) Then
                PlayerAttackNpc index, I, TempPlayer(index).ProjecTile(PlayerProjectile).Damage
                ClearProjectile index, PlayerProjectile
                Exit Sub
            Else
                ClearProjectile index, PlayerProjectile
                Exit Sub
            End If
        End If
    Next
    
    ' hit a block
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
        ' hit a block, clear it.
        ClearProjectile index, PlayerProjectile
        Exit Sub
    End If
    
End Sub
