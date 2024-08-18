Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long, FrameTime As Long
    Dim tmr25 As Long, tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdateMapSpawnItems As Long, LastUpdatePlayerVitals As Long, LastUpdatePlayerRequests As Long

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
        ElapsedTime = Tick - FrameTime
        FrameTime = Tick
        
        If Tick > tmr25 Then
            For i = 1 To Player_HighIndex
                If isPlaying(i) Then
                    ' check if they've completed casting, and if so set the actual spell going
                    If TempPlayer(i).spellBuffer.Spell > 0 Then
                        If GetTickCount > TempPlayer(i).spellBuffer.Timer + (Spell(Player(i).Spell(TempPlayer(i).spellBuffer.Spell)).CastTime * 1000) Then
                            CastSpell i, TempPlayer(i).spellBuffer.Spell, TempPlayer(i).spellBuffer.target, TempPlayer(i).spellBuffer.tType
                            TempPlayer(i).spellBuffer.Spell = 0
                            TempPlayer(i).spellBuffer.Timer = 0
                            TempPlayer(i).spellBuffer.target = 0
                            TempPlayer(i).spellBuffer.tType = 0
                        End If
                    End If
                    ' check if need to turn off stunned
                    If TempPlayer(i).StunDuration > 0 Then
                        If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                            TempPlayer(i).StunDuration = 0
                            TempPlayer(i).StunTimer = 0
                            SendStunned i
                        End If
                    End If
                    ' check regen timer
                    If TempPlayer(i).stopRegen Then
                        If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                            TempPlayer(i).stopRegen = False
                            TempPlayer(i).stopRegenTimer = 0
                        End If
                    End If
                    ' HoT and DoT logic
                    For x = 1 To MAX_DOTS
                        HandleDoT_Player i, x
                        HandleHoT_Player i, x
                    Next
                End If
                
                UpdateEventLogic
            Next
            frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
            tmr25 = GetTickCount + 25
        End If

        ' Check for disconnections every half second
        If Tick > tmr500 Then
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            UpdateMapLogic
            tmr500 = GetTickCount + 500
        End If

        If Tick > tmr1000 Then
            If isShuttingDown Then
                Call HandleShutdown
            End If
            tmr1000 = GetTickCount + 1000
        End If
        
        'LEFTOFF
        If Options.Projectiles = 1 Then
            For i = 1 To Player_HighIndex
                If isPlaying(i) Then
                    For x = 1 To MAX_PLAYER_PROJECTILES
                        If TempPlayer(i).ProjecTile(x).Pic > 0 Then
                            ' handle the projec tile
                            HandleProjecTile i, x
                        End If
                    Next
                End If
            Next
        End If

        ' Checks to update player vitals every 5 seconds - Can be tweaked
        If Tick > LastUpdatePlayerVitals Then
            UpdatePlayerVitals
            LastUpdatePlayerVitals = GetTickCount + 5000
        End If

        ' Checks to spawn map items every 5 minutes - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 300000
        End If

        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If
        
        ' Checks to remove a point from players request count every 5 minutes - Can be tweaked
        If Tick > LastUpdatePlayerRequests Then
            UpdatePlayerFriendRequests
            LastUpdatePlayerRequests = GetTickCount + 300000
        End If
   Dim Var1 As String
   
        ' Checks to see if a player's requests sent should be reset
        
        'Handles Guild Invites
        For i = 1 To Player_HighIndex
            If isPlaying(i) Then
                If TempPlayer(i).tmpGuildInviteSlot > 0 Then
                    If Tick > TempPlayer(i).tmpGuildInviteTimer Then
                        If GuildData(TempPlayer(i).tmpGuildInviteSlot).In_Use = True Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M1")
                            PlayerMsg i, Var1 & GuildData(TempPlayer(i).tmpGuildInviteSlot).Guild_Name & ".", BrightRed
                            TempPlayer(i).tmpGuildInviteSlot = 0    '[SERVIDOR]M1
                            TempPlayer(i).tmpGuildInviteTimer = 0
                        Else
                            'Just remove this guild has been unloaded
                            TempPlayer(i).tmpGuildInviteSlot = 0
                            TempPlayer(i).tmpGuildInviteTimer = 0
                        End If
                    End If
                End If
            End If
        Next i

        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
    Loop
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim y As Long

    ' ///////////////////////////////////////////
    ' // This is used for respawning map items //
    ' ///////////////////////////////////////////
    For y = 1 To MAX_MAPS

        ' Make sure no one is on the map when it respawns
        If Not PlayersOnMap(y) Then

            ' Clear out unnecessary junk
            For x = 1 To MAX_MAP_ITEMS
                Call ClearMapItem(x, y)
            Next

            ' Spawn the items
            Call SpawnMapItems(y)
            Call SendMapItemsToAll(y)
        End If

        DoEvents
    Next

End Sub

Private Sub UpdateMapLogic()
On Error GoTo error:
    Dim i As Long, x As Long, mapnum As Long, n As Long, x1 As Long, y1 As Long
    Dim TickCount As Long, Damage As Long, DistanceX As Long, DistanceY As Long, npcNum As Long
    Dim target As Long, targetType As Byte, didwalk As Boolean, buffer As clsBuffer, Resource_index As Long
    Dim targetX As Long, targetY As Long, target_verify As Boolean
    Dim SecondsToSpawn As Long, SetSpawnTime As Boolean

    For mapnum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(mapnum, i).num > 0 Then
                If MapItem(mapnum, i).playerName <> vbNullString Then
                    ' make item public?
                    If MapItem(mapnum, i).playerTimer < GetTickCount Then
                        ' make it public
                        MapItem(mapnum, i).playerName = vbNullString
                        MapItem(mapnum, i).playerTimer = 0
                        ' send updates to everyone
                        SendMapItemsToAll mapnum
                    End If
                    ' despawn item?
                    If MapItem(mapnum, i).canDespawn Then
                        If MapItem(mapnum, i).despawnTimer < GetTickCount Then
                            ' despawn it
                            ClearMapItem i, mapnum
                            ' send updates to everyone
                            SendMapItemsToAll mapnum
                        End If
                    End If
                End If
            End If
        Next
        
        '  Close the doors
        If TickCount > temptile(mapnum).DoorTimer + 5000 Then
            For x1 = 0 To Map(mapnum).MaxX
                For y1 = 0 To Map(mapnum).MaxY
                    If Map(mapnum).Tile(x1, y1).Type = TILE_TYPE_KEY And temptile(mapnum).DoorOpen(x1, y1) = YES Then
                        temptile(mapnum).DoorOpen(x1, y1) = NO
                        SendMapKeyToMap mapnum, x1, y1, 0
                    End If
                Next
            Next
        End If
        
        ' check for DoTs + hots
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(mapnum).NPC(i).num > 0 Then
                For x = 1 To MAX_DOTS
                    HandleDoT_Npc mapnum, i, x
                    HandleHoT_Npc mapnum, i, x
                Next
            End If
        Next

        ' Respawning Resources
        If ResourceCache(mapnum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(mapnum).Resource_Count
                Resource_index = Map(mapnum).Tile(ResourceCache(mapnum).ResourceData(i).x, ResourceCache(mapnum).ResourceData(i).y).Data1

                If Resource_index > 0 Then
                    If ResourceCache(mapnum).ResourceData(i).ResourceState = 1 Or ResourceCache(mapnum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(mapnum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(mapnum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(mapnum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root/check if random
                            If Resource(Resource_index).HPRand = 0 Then
                                ResourceCache(mapnum).ResourceData(i).HPSetTo = Resource(Resource_index).health
                                ResourceCache(mapnum).ResourceData(i).cur_health = Resource(Resource_index).health
                            Else
                                ResourceCache(mapnum).ResourceData(i).HPSetTo = rand(Resource(Resource_index).healthmin, Resource(Resource_index).health)
                                ResourceCache(mapnum).ResourceData(i).cur_health = ResourceCache(mapnum).ResourceData(i).HPSetTo
                            End If
                            SendResourceCacheToMap mapnum, i
                        End If
                    End If
                End If
            Next
        End If

        If PlayersOnMap(mapnum) = YES Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPCS
                npcNum = MapNpc(mapnum).NPC(x).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_GUARD Then
                    
                        ' make sure it's not stunned
                        If Not MapNpc(mapnum).NPC(x).StunDuration > 0 Then
    
                            For i = 1 To Player_HighIndex
                                If isPlaying(i) Then
                                    If GetPlayerMap(i) = mapnum And MapNpc(mapnum).NPC(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = NPC(npcNum).Range
                                        DistanceX = MapNpc(mapnum).NPC(x).x - GetPlayerX(i)
                                        DistanceY = MapNpc(mapnum).NPC(x).y - GetPlayerY(i)
    
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
    
                                        ' Are they in range?  if so GET'M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If NPC(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or GetPlayerPK(i) = YES Then
                                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                                    Call SendChatBubble(mapnum, x, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                                                End If
                                                MapNpc(mapnum).NPC(x).targetType = 1 ' player
                                                MapNpc(mapnum).NPC(x).target = i
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                
                target_verify = False

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then
                    If MapNpc(mapnum).NPC(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(mapnum).NPC(x).StunTimer + (MapNpc(mapnum).NPC(x).StunDuration * 1000) Then
                            MapNpc(mapnum).NPC(x).StunDuration = 0
                            MapNpc(mapnum).NPC(x).StunTimer = 0
                        End If
                    Else
                            
                        target = MapNpc(mapnum).NPC(x).target
                        targetType = MapNpc(mapnum).NPC(x).targetType
    
                        ' Check to see if its time for the npc to walk
                        If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                        
                            If targetType = 1 Then ' player
    
                                ' Check to see if we are following a player or not
                                If target > 0 Then
        
                                    ' Check if the player is even playing, if so follow'm
                                    If isPlaying(target) And GetPlayerMap(target) = mapnum Then
                                        didwalk = False
                                        target_verify = True
                                        targetY = GetPlayerY(target)
                                        targetX = GetPlayerX(target)
                                    Else
                                        MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapnum).NPC(x).target = 0
                                    End If
                                End If
                            
                            ElseIf targetType = 2 Then 'npc
                                
                                If target > 0 Then
                                    
                                    If MapNpc(mapnum).NPC(target).num > 0 Then
                                        didwalk = False
                                        target_verify = True
                                        targetY = MapNpc(mapnum).NPC(target).y
                                        targetX = MapNpc(mapnum).NPC(target).x
                                    Else
                                        MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                                        MapNpc(mapnum).NPC(x).target = 0
                                    End If
                                End If
                            End If
                            
                            If target_verify Then
                                'Gonna make the npcs smarter.. Implementing a pathfinding algorithm.. we shall see what happens.
                                If IsOneBlockAway(targetX, targetY, CLng(MapNpc(mapnum).NPC(x).x), CLng(MapNpc(mapnum).NPC(x).y)) = False Then
                                    If PathfindingType = 1 Then
                                        i = Int(Rnd * 5)
            
                                        ' Lets move the npc
                                        Select Case i
                                            Case 0
            
                                                ' Up
                                                If MapNpc(mapnum).NPC(x).y > targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_UP) Then
                                                        Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(mapnum).NPC(x).y < targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                        Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(mapnum).NPC(x).x > targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                        Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(mapnum).NPC(x).x < targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                        Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 1
            
                                                ' Right
                                                If MapNpc(mapnum).NPC(x).x < targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                        Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(mapnum).NPC(x).x > targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                        Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(mapnum).NPC(x).y < targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                        Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(mapnum).NPC(x).y > targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_UP) Then
                                                        Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 2
            
                                                ' Down
                                                If MapNpc(mapnum).NPC(x).y < targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                        Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(mapnum).NPC(x).y > targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_UP) Then
                                                        Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(mapnum).NPC(x).x < targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                        Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Left
                                                If MapNpc(mapnum).NPC(x).x > targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                        Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                            Case 3
            
                                                ' Left
                                                If MapNpc(mapnum).NPC(x).x > targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_LEFT) Then
                                                        Call NpcMove(mapnum, x, DIR_LEFT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Right
                                                If MapNpc(mapnum).NPC(x).x < targetX And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_RIGHT) Then
                                                        Call NpcMove(mapnum, x, DIR_RIGHT, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Up
                                                If MapNpc(mapnum).NPC(x).y > targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_UP) Then
                                                        Call NpcMove(mapnum, x, DIR_UP, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                                ' Down
                                                If MapNpc(mapnum).NPC(x).y < targetY And Not didwalk Then
                                                    If CanNpcMove(mapnum, x, DIR_DOWN) Then
                                                        Call NpcMove(mapnum, x, DIR_DOWN, MOVING_WALKING)
                                                        didwalk = True
                                                    End If
                                                End If
            
                                        End Select
            
                                        ' Check if we can't move and if Target is behind something and if we can just switch dirs
                                        If Not didwalk Then
                                            If MapNpc(mapnum).NPC(x).x - 1 = targetX And MapNpc(mapnum).NPC(x).y = targetY Then
                                                If MapNpc(mapnum).NPC(x).Dir <> DIR_LEFT Then
                                                    Call NpcDir(mapnum, x, DIR_LEFT)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(mapnum).NPC(x).x + 1 = targetX And MapNpc(mapnum).NPC(x).y = targetY Then
                                                If MapNpc(mapnum).NPC(x).Dir <> DIR_RIGHT Then
                                                    Call NpcDir(mapnum, x, DIR_RIGHT)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(mapnum).NPC(x).x = targetX And MapNpc(mapnum).NPC(x).y - 1 = targetY Then
                                                If MapNpc(mapnum).NPC(x).Dir <> DIR_UP Then
                                                    Call NpcDir(mapnum, x, DIR_UP)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            If MapNpc(mapnum).NPC(x).x = targetX And MapNpc(mapnum).NPC(x).y + 1 = targetY Then
                                                If MapNpc(mapnum).NPC(x).Dir <> DIR_DOWN Then
                                                    Call NpcDir(mapnum, x, DIR_DOWN)
                                                End If
            
                                                didwalk = True
                                            End If
            
                                            ' We could not move so Target must be behind something, walk randomly.
                                            If Not didwalk Then
                                                i = Int(Rnd * 2)
            
                                                If i = 1 Then
                                                    i = Int(Rnd * 4)
            
                                                    If CanNpcMove(mapnum, x, i) Then
                                                        Call NpcMove(mapnum, x, i, MOVING_WALKING)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        i = FindNpcPath(mapnum, x, targetX, targetY)
                                        If i < 4 Then 'Returned an answer. Move the NPC
                                            If CanNpcMove(mapnum, x, i) Then
                                                NpcMove mapnum, x, i, MOVING_WALKING
                                            End If
                                        Else 'No good path found. Move randomly
                                            i = Int(Rnd * 4)
                                            If i = 1 Then
                                                i = Int(Rnd * 4)
                
                                                If CanNpcMove(mapnum, x, i) Then
                                                    Call NpcMove(mapnum, x, i, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    Call NpcDir(mapnum, x, GetNpcDir(targetX, targetY, CLng(MapNpc(mapnum).NPC(x).x), CLng(MapNpc(mapnum).NPC(x).y)))
                                End If
                            Else
                                i = Int(Rnd * 4)
    
                                If i = 1 Then
                                    i = Int(Rnd * 4)
    
                                    If CanNpcMove(mapnum, x, i) Then
                                        Call NpcMove(mapnum, x, i, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(mapnum).NPC(x) > 0 And MapNpc(mapnum).NPC(x).num > 0 Then
                    target = MapNpc(mapnum).NPC(x).target
                    targetType = MapNpc(mapnum).NPC(x).targetType

                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                    
                        If targetType = 1 Then ' player

                            ' Is the target playing and on the same map?
                            If isPlaying(target) And GetPlayerMap(target) = mapnum Then
                                TryNpcAttackPlayer x, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(mapnum).NPC(x).target = 0
                                MapNpc(mapnum).NPC(x).targetType = 0 ' clear
                            End If
                        Else
                            ' lol no npc combat :(
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If Not MapNpc(mapnum).NPC(x).stopRegen Then
                    If MapNpc(mapnum).NPC(x).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(mapnum).NPC(x).Vital(Vitals.HP) > 0 Then
                            MapNpc(mapnum).NPC(x).Vital(Vitals.HP) = MapNpc(mapnum).NPC(x).Vital(Vitals.HP) + GetNpcVitalRegen(npcNum, Vitals.HP)
    
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(mapnum).NPC(x).Vital(Vitals.HP) > MapNpc(mapnum).NPC(x).HPSetTo Then
                                If MapNpc(mapnum).NPC(x).HPSetTo > 0 Then
                                    MapNpc(mapnum).NPC(x).Vital(Vitals.HP) = MapNpc(mapnum).NPC(x).HPSetTo
                                Else
                                    MapNpc(mapnum).NPC(x).Vital(Vitals.HP) = GetNpcMaxVital(x, Vitals.HP)
                                End If
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).STR > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(mapnum).NPC(x).num = 0 And Map(mapnum).NPC(x) > 0 Then
                    SecondsToSpawn = NPC(Map(mapnum).NPC(x)).SpawnSecs * 1000
                    If NPC(Map(mapnum).NPC(x)).RndSpawn = 1 Then
                        SecondsToSpawn = rand(NPC(Map(mapnum).NPC(x)).SpawnSecsMin, NPC(Map(mapnum).NPC(x)).SpawnSecs) * 1000
                    End If
                    If TickCount > MapNpc(mapnum).NPC(x).SpawnWait + SecondsToSpawn Then
                        Call SpawnNpc(x, mapnum)
                    End If
                End If

            Next

        End If

        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If
error:
End Sub

Private Sub UpdatePlayerFriendRequests()
Dim i As Long
    For i = 1 To Player_HighIndex
        If isPlaying(i) Then
            If GetPlayerFriendRequests(i) > 0 Then
                Call SetPlayerFriendRequests(i, -1)
            End If
        End If
    Next i
End Sub


Private Sub UpdatePlayerVitals()
Dim i As Long
    For i = 1 To Player_HighIndex
        If isPlaying(i) Then
            If Not TempPlayer(i).stopRegen Then
                If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                    Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                    Call SendVital(i, Vitals.HP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
    
                If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                    Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                    Call SendVital(i, Vitals.MP)
                    ' send vitals to party if in one
                    If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                End If
            End If
        End If
    Next
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    If TotalOnlinePlayers > 0 Then
        Dim Var1 As String
        
        
Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s11")
        Call TextAdd(Var1) '[LOGS]s11

        For i = 1 To Player_HighIndex

            If isPlaying(i) Then
                Call SavePlayer(i)
                Call SaveBank(i)
            End If

            DoEvents
        Next

    End If

End Sub

Private Sub HandleShutdown()
Dim Var1 As String

Dim var2 As String


    If Secs <= 0 Then Secs = 30
    If Secs Mod 5 = 0 Or Secs <= 5 Then     '[SERVIDOR]M2 y M3
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M2")
                        var2 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M3")
        Call GlobalMsg(Var1 & Secs & var2, BrightBlue)
        Call TextAdd(Var1 & Secs & var2)
    End If                                '[SERVIDOR]M2 y M3

    Secs = Secs - 1

    If Secs <= 0 Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M4")
        Call GlobalMsg(Var1, BrightRed) '[SERVIDOR]M4
        Call DestroyServer
    End If

End Sub

Function CanEventMoveTowardsPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long
Dim Path() As Vector, LastX As Long, LastY As Long, did As Boolean
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 4 is not a valid direction so we assume fail unless told otherwise.
    CanEventMoveTowardsPlayer = 4
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(mapnum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    'Add option for pathfinding to random guessing option.
    
    If PathfindingType = 1 Then
        i = Int(Rnd * 5)
        didwalk = False
        
        ' Lets move the event
        Select Case i
            Case 0
        
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 1
            
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 2
            
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
        
            Case 3
            
                ' Left
                If x1 > x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                        CanEventMoveTowardsPlayer = DIR_LEFT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Right
                If x1 < x And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                        CanEventMoveTowardsPlayer = DIR_RIGHT
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Up
                If y1 > y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                        CanEventMoveTowardsPlayer = DIR_UP
                        Exit Function
                        didwalk = True
                    End If
                End If
                
                ' Down
                If y1 < y And Not didwalk Then
                    If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                        CanEventMoveTowardsPlayer = DIR_DOWN
                        Exit Function
                        didwalk = True
                    End If
                End If
        End Select
        CanEventMoveTowardsPlayer = Random(0, 3)
    ElseIf PathfindingType = 2 Then
        'Initialization phase
        tim = 0
        sX = x1
        sY = y1
        FX = x
        FY = y
        
        ReDim pos(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
        
        'CacheMapBlocks mapnum
        
        pos = MapBlocks(mapnum).Blocks
        
        For i = 1 To TempPlayer(playerID).EventMap.CurrentEvents
            If TempPlayer(playerID).EventMap.EventPages(i).Visible Then
                If TempPlayer(playerID).EventMap.EventPages(i).WalkThrough = 1 Then
                    pos(TempPlayer(playerID).EventMap.EventPages(i).x, TempPlayer(playerID).EventMap.EventPages(i).y) = 9
                End If
            End If
        Next
        
        pos(sX, sY) = 100 + tim
        pos(FX, FY) = 2
        
        'reset reachable
        reachable = False
        
        'Do while reachable is false... if its set true in progress, we jump out
        'If the path is decided unreachable in process, we will use exit sub. Not proper,
        'but faster ;-)
        Do While reachable = False
            'we loop through all squares
            For j = 0 To Map(mapnum).MaxY
                For i = 0 To Map(mapnum).MaxX
                    'If j = 10 And i = 0 Then MsgBox "hi!"
                    'If they are to be extended, the pointer TIM is on them
                    If pos(i, j) = 100 + tim Then
                    'The part is to be extended, so do it
                        'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                        'because then we get error... If the square is on side, we dont test for this one!
                        If i < Map(mapnum).MaxX Then
                            'If there isnt a wall, or any other... thing
                            If pos(i + 1, j) = 0 Then
                                'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                                'It will exapand that square too! This is crucial part of the program
                                pos(i + 1, j) = 100 + tim + 1
                            ElseIf pos(i + 1, j) = 2 Then
                                'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                                reachable = True
                            End If
                        End If
                    
                        'This is the same as the last one, as i said a lot of copy paste work and editing that
                        'This is simply another side that we have to test for... so instead of i+1 we have i-1
                        'Its actually pretty same then... I wont comment it therefore, because its only repeating
                        'same thing with minor changes to check sides
                        If i > 0 Then
                            If pos((i - 1), j) = 0 Then
                                pos(i - 1, j) = 100 + tim + 1
                            ElseIf pos(i - 1, j) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j < Map(mapnum).MaxY Then
                            If pos(i, j + 1) = 0 Then
                                pos(i, j + 1) = 100 + tim + 1
                            ElseIf pos(i, j + 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    
                        If j > 0 Then
                            If pos(i, j - 1) = 0 Then
                                pos(i, j - 1) = 100 + tim + 1
                            ElseIf pos(i, j - 1) = 2 Then
                                reachable = True
                            End If
                        End If
                    End If
                    DoEvents
                Next i
            Next j
            
            'If the reachable is STILL false, then
            If reachable = False Then
                'reset sum
                Sum = 0
                For j = 0 To Map(mapnum).MaxY
                    For i = 0 To Map(mapnum).MaxX
                    'we add up ALL the squares
                    Sum = Sum + pos(i, j)
                    Next i
                Next j
                
                'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
                'sum to lastsum
                If Sum = LastSum Then
                    CanEventMoveTowardsPlayer = 4
                    Exit Function
                Else
                    LastSum = Sum
                End If
            End If
            
            'we increase the pointer to point to the next squares to be expanded
            tim = tim + 1
        Loop
        
        'We work backwards to find the way...
        LastX = FX
        LastY = FY
        
        ReDim Path(tim + 1)
        
        'The following code may be a little bit confusing but ill try my best to explain it.
        'We are working backwards to find ONE of the shortest ways back to Start.
        'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
        'how LastX and LasY change
        Do While LastX <> sX Or LastY <> sY
            'We decrease tim by one, and then we are finding any adjacent square to the final one, that
            'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
            'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
            'that value. When we find it, we just color it yellow as for the solution
            tim = tim - 1
            'reset did to false
            did = False
            
            'If we arent on edge
            If LastX < Map(mapnum).MaxX Then
                'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
                If pos(LastX + 1, LastY) = 100 + tim Then
                    'if it, then make it yellow, and change did to true
                    LastX = LastX + 1
                    did = True
                End If
            End If
            
            'This will then only work if the previous part didnt execute, and did is still false. THen
            'we want to check another square, the on left. Is it a tim-1 one ?
            If did = False Then
                If LastX > 0 Then
                    If pos(LastX - 1, LastY) = 100 + tim Then
                        LastX = LastX - 1
                        did = True
                    End If
                End If
            End If
            
            'We check the one below it
            If did = False Then
                If LastY < Map(mapnum).MaxY Then
                    If pos(LastX, LastY + 1) = 100 + tim Then
                        LastY = LastY + 1
                        did = True
                    End If
                End If
            End If
            
            'And above it. One of these have to be it, since we have found the solution, we know that already
            'there is a way back.
            If did = False Then
                If LastY > 0 Then
                    If pos(LastX, LastY - 1) = 100 + tim Then
                        LastY = LastY - 1
                    End If
                End If
            End If
            
            Path(tim).x = LastX
            Path(tim).y = LastY
            
            'Now we loop back and decrease tim, and look for the next square with lower value
            DoEvents
        Loop
        
        'Ok we got a path. Now, lets look at the first step and see what direction we should take.
        If Path(1).x > LastX Then
            CanEventMoveTowardsPlayer = DIR_RIGHT
        ElseIf Path(1).y > LastY Then
            CanEventMoveTowardsPlayer = DIR_DOWN
        ElseIf Path(1).y < LastY Then
            CanEventMoveTowardsPlayer = DIR_UP
        ElseIf Path(1).x < LastX Then
            CanEventMoveTowardsPlayer = DIR_LEFT
        End If
        
    End If
End Function

Function CanEventMoveAwayFromPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    CanEventMoveAwayFromPlayer = 5
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    WalkThrough = Map(mapnum).Events(TempPlayer(playerID).EventMap.EventPages(eventID).eventID).Pages(TempPlayer(playerID).EventMap.EventPages(eventID).pageID).WalkThrough
    
    i = Int(Rnd * 5)
    didwalk = False
    
    ' Lets move the event
    Select Case i
        Case 0
    
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 1
        
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 2
        
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
    
        Case 3
        
            ' Left
            If x1 > x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_RIGHT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_RIGHT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Right
            If x1 < x And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_LEFT, False) Then
                    CanEventMoveAwayFromPlayer = DIR_LEFT
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Up
            If y1 > y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_DOWN, False) Then
                    CanEventMoveAwayFromPlayer = DIR_DOWN
                    Exit Function
                    didwalk = True
                End If
            End If
            
            ' Down
            If y1 < y And Not didwalk Then
                If CanEventMove(playerID, mapnum, x1, y1, eventID, WalkThrough, DIR_UP, False) Then
                    CanEventMoveAwayFromPlayer = DIR_UP
                    Exit Function
                    didwalk = True
                End If
            End If
    
        End Select
        
        CanEventMoveAwayFromPlayer = Random(0, 3)
End Function

Function GetDirToPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_DOWN
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirToPlayer = i
    
End Function

Function GetDirAwayFromPlayer(playerID As Long, mapnum As Long, eventID As Long) As Long
Dim i As Long, x As Long, y As Long, x1 As Long, y1 As Long, didwalk As Boolean, WalkThrough As Long, distance As Long
    'This does not work for global events so this MUST be a player one....
    'This Event returns a direction, 5 is not a valid direction so we assume fail unless otherwise told.
    If playerID <= 0 Or playerID > Player_HighIndex Then Exit Function
    If mapnum <= 0 Or mapnum > MAX_MAPS Then Exit Function
    If eventID <= 0 Or eventID > TempPlayer(playerID).EventMap.CurrentEvents Then Exit Function
    
    x = GetPlayerX(playerID)
    y = GetPlayerY(playerID)
    x1 = TempPlayer(playerID).EventMap.EventPages(eventID).x
    y1 = TempPlayer(playerID).EventMap.EventPages(eventID).y
    
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_LEFT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_RIGHT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_UP
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_DOWN
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetDirAwayFromPlayer = i
End Function

Function FindNpcPath(mapnum As Long, mapNpcNum As Long, targetX As Long, targetY As Long) As Long
Dim tim As Long, sX As Long, sY As Long, pos() As Long, reachable As Boolean, x As Long, y As Long, j As Long, LastSum As Long, Sum As Long, FX As Long, FY As Long, i As Long
Dim Path() As Vector, LastX As Long, LastY As Long, did As Boolean

'Initialization phase
tim = 0
sX = MapNpc(mapnum).NPC(mapNpcNum).x
sY = MapNpc(mapnum).NPC(mapNpcNum).y
FX = targetX
FY = targetY

ReDim pos(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
pos = MapBlocks(mapnum).Blocks

pos(sX, sY) = 100 + tim
pos(FX, FY) = 2

'reset reachable
reachable = False

'Do while reachable is false... if its set true in progress, we jump out
'If the path is decided unreachable in process, we will use exit sub. Not proper,
'but faster ;-)
Do While reachable = False
    'we loop through all squares
    For j = 0 To Map(mapnum).MaxY
        For i = 0 To Map(mapnum).MaxX
            'If j = 10 And i = 0 Then MsgBox "hi!"
            'If they are to be extended, the pointer TIM is on them
            If pos(i, j) = 100 + tim Then
            'The part is to be extended, so do it
                'We have to make sure that there is a pos(i+1,j) BEFORE we actually use it,
                'because then we get error... If the square is on side, we dont test for this one!
                If i < Map(mapnum).MaxX Then
                    'If there isnt a wall, or any other... thing
                    If pos(i + 1, j) = 0 Then
                        'Expand it, and make its pos equal to tim+1, so the next time we make this loop,
                        'It will exapand that square too! This is crucial part of the program
                        pos(i + 1, j) = 100 + tim + 1
                    ElseIf pos(i + 1, j) = 2 Then
                        'If the position is no 0 but its 2 (FINISH) then Reachable = true!!! We found end
                        reachable = True
                    End If
                End If
            
                'This is the same as the last one, as i said a lot of copy paste work and editing that
                'This is simply another side that we have to test for... so instead of i+1 we have i-1
                'Its actually pretty same then... I wont comment it therefore, because its only repeating
                'same thing with minor changes to check sides
                If i > 0 Then
                    If pos((i - 1), j) = 0 Then
                        pos(i - 1, j) = 100 + tim + 1
                    ElseIf pos(i - 1, j) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j < Map(mapnum).MaxY Then
                    If pos(i, j + 1) = 0 Then
                        pos(i, j + 1) = 100 + tim + 1
                    ElseIf pos(i, j + 1) = 2 Then
                        reachable = True
                    End If
                End If
            
                If j > 0 Then
                    If pos(i, j - 1) = 0 Then
                        pos(i, j - 1) = 100 + tim + 1
                    ElseIf pos(i, j - 1) = 2 Then
                        reachable = True
                    End If
                End If
            End If
            DoEvents
        Next i
    Next j
    
    'If the reachable is STILL false, then
    If reachable = False Then
        'reset sum
        Sum = 0
        For j = 0 To Map(mapnum).MaxY
            For i = 0 To Map(mapnum).MaxX
            'we add up ALL the squares
            Sum = Sum + pos(i, j)
            Next i
        Next j
        
        'Now if the sum is euqal to the last sum, its not reachable, if it isnt, then we store
        'sum to lastsum
        If Sum = LastSum Then
            FindNpcPath = 4
            Exit Function
        Else
            LastSum = Sum
        End If
    End If
    
    'we increase the pointer to point to the next squares to be expanded
    tim = tim + 1
Loop

'We work backwards to find the way...
LastX = FX
LastY = FY

ReDim Path(tim + 1)

'The following code may be a little bit confusing but ill try my best to explain it.
'We are working backwards to find ONE of the shortest ways back to Start.
'So we repeat the loop until the LastX and LastY arent in start. Look in the code to see
'how LastX and LasY change
Do While LastX <> sX Or LastY <> sY
    'We decrease tim by one, and then we are finding any adjacent square to the final one, that
    'has that value. So lets say the tim would be 5, because it takes 5 steps to get to the target.
    'Now everytime we decrease that, so we make it 4, and we look for any adjacent square that has
    'that value. When we find it, we just color it yellow as for the solution
    tim = tim - 1
    'reset did to false
    did = False
    
    'If we arent on edge
    If LastX < Map(mapnum).MaxX Then
        'check the square on the right of the solution. Is it a tim-1 one? or just a blank one
        If pos(LastX + 1, LastY) = 100 + tim Then
            'if it, then make it yellow, and change did to true
            LastX = LastX + 1
            did = True
        End If
    End If
    
    'This will then only work if the previous part didnt execute, and did is still false. THen
    'we want to check another square, the on left. Is it a tim-1 one ?
    If did = False Then
        If LastX > 0 Then
            If pos(LastX - 1, LastY) = 100 + tim Then
                LastX = LastX - 1
                did = True
            End If
        End If
    End If
    
    'We check the one below it
    If did = False Then
        If LastY < Map(mapnum).MaxY Then
            If pos(LastX, LastY + 1) = 100 + tim Then
                LastY = LastY + 1
                did = True
            End If
        End If
    End If
    
    'And above it. One of these have to be it, since we have found the solution, we know that already
    'there is a way back.
    If did = False Then
        If LastY > 0 Then
            If pos(LastX, LastY - 1) = 100 + tim Then
                LastY = LastY - 1
            End If
        End If
    End If
    
    Path(tim).x = LastX
    Path(tim).y = LastY
    
    'Now we loop back and decrease tim, and look for the next square with lower value
    DoEvents
Loop

'Ok we got a path. Now, lets look at the first step and see what direction we should take.
If Path(1).x > LastX Then
    FindNpcPath = DIR_RIGHT
ElseIf Path(1).y > LastY Then
    FindNpcPath = DIR_DOWN
ElseIf Path(1).y < LastY Then
    FindNpcPath = DIR_UP
ElseIf Path(1).x < LastX Then
    FindNpcPath = DIR_LEFT
End If

End Function

Public Sub CacheMapBlocks(mapnum As Long)
Dim x As Long, y As Long
    ReDim MapBlocks(mapnum).Blocks(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            If NpcTileIsOpen(mapnum, x, y) = False Then
                MapBlocks(mapnum).Blocks(x, y) = 9
            End If
        Next
    Next
     
End Sub

Public Sub UpdateMapBlock(mapnum, x, y, blocked As Boolean)
On Error GoTo error:
    If blocked Then
        MapBlocks(mapnum).Blocks(x, y) = 9
    Else
        MapBlocks(mapnum).Blocks(x, y) = 0
    End If
error:
End Sub

Function IsOneBlockAway(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Boolean
    If x1 = x2 Then
        If y1 = y2 - 1 Or y1 = y2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    ElseIf y1 = y2 Then
        If x1 = x2 - 1 Or x1 = x2 + 1 Then
            IsOneBlockAway = True
        Else
            IsOneBlockAway = False
        End If
    Else
        IsOneBlockAway = False
    End If
End Function

Function GetNpcDir(x As Long, y As Long, x1 As Long, y1 As Long) As Long
    Dim i As Long, distance As Long
    
    i = DIR_RIGHT
    
    If x - x1 > 0 Then
        If x - x1 > distance Then
            i = DIR_RIGHT
            distance = x - x1
        End If
    ElseIf x - x1 < 0 Then
        If ((x - x1) * -1) > distance Then
            i = DIR_LEFT
            distance = ((x - x1) * -1)
        End If
    End If
    
    If y - y1 > 0 Then
        If y - y1 > distance Then
            i = DIR_DOWN
            distance = y - y1
        End If
    ElseIf y - y1 < 0 Then
        If ((y - y1) * -1) > distance Then
            i = DIR_UP
            distance = ((y - y1) * -1)
        End If
    End If
    
    GetNpcDir = i
        
End Function
