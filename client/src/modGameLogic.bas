Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
Dim FrameTime As Long
Dim tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim I As Long
Dim WalkTimer As Long
Dim tmr25 As Long
Dim tmr100 As Long
Dim tmr10000 As Long
Dim tmr500, Fadetmr As Long
Dim fogtmr As Long
Dim chatTmr As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' *** Start GameLoop ***
    Do While InGame
        tick = GetTickCount                            ' Set the inital tick
        ElapsedTime = tick - FrameTime                 ' Set the time difference for time-based movement
        FrameTime = tick                               ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then


            
            ' check ping
            Call GetPing
            Call DrawPing
            tmr10000 = tick + 10000
        End If

        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            If GetForegroundWindow() = frmMain.hwnd Or GetForegroundWindow() = frmEditor_Events.hwnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For I = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(I) > 0 Then
                        If SpellCD(I) > 0 Then
                            If SpellCD(I) + (Spell(PlayerSpells(I)).CDTime * 1000) < tick Then
                                SpellCD(I) = 0
                            End If
                        End If
                    End If
                Next
            End If
            
            ' check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                    SpellBuffer = 0
                    SpellBufferTimer = 0
                End If
            End If

            If CanMoveNow Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack   ' Check to see if player is trying to attack
                
            End If

            ' Change map animation every 250 milliseconds
            If MapAnimTimer < tick Then
                MapAnim = Not MapAnim
                MapAnimTimer = tick + 250
            End If
            
            ' Update inv animation
            If numitems > 0 Then
                If tmr100 < tick Then
                    DrawAnimatedInvItems
                    tmr100 = tick + 100
                End If
            End If
            
            For I = 1 To MAX_BYTE
                CheckAnimInstance I
            Next
            
            If chatTmr < tick Then
                If ChatButtonUp Then
                    ScrollChatBox 0
                End If
                If ChatButtonDown Then
                    ScrollChatBox 1
                End If
                chatTmr = tick + 50
            End If
            
            tmr25 = tick + 25
        End If
        
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then

            For I = 1 To Player_HighIndex
                If IsPlaying(I) Then
                    Call ProcessMovement(I)
                End If
            Next I

            ' Process npc movements (actually move them)
            For I = 1 To Npc_HighIndex
                If Map.NPC(I) > 0 Then
                    Call ProcessNpcMovement(I)
                End If
            Next I
            
            If Map.CurrentEvents > 0 Then
                For I = 1 To Map.CurrentEvents
                    Call ProcessEventMovement(I)
                Next I
            End If

            WalkTimer = tick + 30 ' edit this value to change WalkTimer
        End If
        
        ' fog scrolling
        If fogtmr < tick Then
            If CurrentFogSpeed > 0 Then
                ' move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                ' reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                fogtmr = tick + 255 - CurrentFogSpeed
            End If
        End If
        
        If tmr500 < tick Then
            ' animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            
            ' animate textbox
            If chatOn Then
                If chatShowLine = "|" Then
                    chatShowLine = vbNullString
                Else
                    chatShowLine = "|"
                End If
            End If
            
            tmr500 = tick + 500
        End If
        
        ProcessWeather
        
        If Fadetmr < tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                        
                    Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                    
                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            Fadetmr = tick + 30
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        Call UpdateSounds
        DoEvents

        ' Lock fps
        If Not FPS_Lock Then
            Do While GetTickCount < tick + 15
                DoEvents
                Sleep 1
            Loop
        End If
        
        ' Calculate fps
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
        
    'Anticheat Invisibilidad y otros Buff EaSee 0.5
If MyIndex <> 0 Then
If InvisibilidadDuracion = 0 And GetPlayerVisible(MyIndex) = 1 And GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
    Call EnviarVisibilidad
End If
'Buff/Debuff
If FuerzaDuracionP = 0 And FBP = True Then
FBP = False
Call RestaurarBuff(1, 1)
End If
If FuerzaDuracionN = 0 And FBN = True Then
FBN = False
Call RestaurarBuff(1, 0)
End If
If DestrezaDuracionP = 0 And DBP = True Then
DBP = False
Call RestaurarBuff(2, 1)
End If
If DestrezaDuracionN = 0 And DBN = True Then
DBN = False
Call RestaurarBuff(2, 0)
End If
If AgilidadDuracionP = 0 And ABP = True Then
ABP = False
Call RestaurarBuff(3, 1)
End If
If AgilidadDuracionN = 0 And ABN = True Then
ABN = False
Call RestaurarBuff(3, 0)
End If
If InteligenciaDuracionP = 0 And IBP = True Then
IBP = False
Call RestaurarBuff(4, 1)
End If
If InteligenciaDuracionN = 0 And IBN = True Then
IBN = False
Call RestaurarBuff(4, 0)
End If
If VoluntadDuracionP = 0 And VBP = True Then
VBP = False
Call RestaurarBuff(5, 1)
End If
If VoluntadDuracionN = 0 And VBN = True Then
VBN = False
Call RestaurarBuff(5, 0)
End If
'Sprite
If SpriteDuracion = 0 And SpriteH = False Then
Call Restaurar_Sprite
End If

End If

    Loop

    frmMain.Visible = False

    If isLogging Then
        isLogging = False
        frmEditor_Character.Visible = False
        frmMenu.Visible = True
        GettingMap = True
        StopMusic
        PlayMusic Options.MenuMusic
    Else
        ' Shutdown the game
        frmLoad.Visible = True
        Call SetStatus("Destroying game data...")
        Call DestroyGame
    End If

    'EaSee 0.5 Fin Invisibilidad
    
    

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long
Dim Clase As Long

Clase = GetPlayerClass(MyIndex)

If VelocidadDuracion < 1 Then 'Prevenimos trucos o hackeos de buff velocidad
VelocidadCaminarH = 0
VelocidadCorrerH = 0
VelocidadBuff = 1
End If

If VelocidadDuracion > 1 And VelocidadBuff = 0 Then
WALK_SPEED = Class(Clase).VCaminar(0) - VelocidadCaminarH
RUN_SPEED = Class(Clase).VCorrer(0) - VelocidadCorrerH
ElseIf VelocidadDuracion > 1 And VelocidadBuff = 1 Then
WALK_SPEED = Class(Clase).VCaminar(0) + VelocidadCaminarH
RUN_SPEED = Class(Clase).VCorrer(0) + VelocidadCorrerH
Else
WALK_SPEED = Class(Clase).VCaminar(0)
RUN_SPEED = Class(Clase).VCorrer(0)
End If

'EaSee 0.5 Fix de sincronizacion

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if player is walking, and if so process moving them over
    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * (WALK_SPEED * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (RUN_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If VXFRAME = False Then
                    If Player(Index).Step = 1 Then
                        Player(Index).Step = 3
                    Else
                        Player(Index).Step = 1
                    End If
                Else
                    If Player(Index).Step = 0 Then
                        Player(Index).Step = 2
                    Else
                        Player(Index).Step = 0
                    End If
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If VXFRAME = False Then
                    If Player(Index).Step = 1 Then
                        Player(Index).Step = 3
                    Else
                        Player(Index).Step = 1
                    End If
                Else
                    If Player(Index).Step = 0 Then
                        Player(Index).Step = 2
                    Else
                        Player(Index).Step = 0
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ProcessMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if NPC is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If MapNpc(MapNpcNum).num > 0 Then
                    MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - ((ElapsedTime / 1000) * (NPC(MapNpc(MapNpcNum).num).speed * SIZE_X))
                    If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
                End If
                
            Case DIR_DOWN
                If MapNpc(MapNpcNum).num > 0 Then
                    MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + ((ElapsedTime / 1000) * (NPC(MapNpc(MapNpcNum).num).speed * SIZE_X))
                    If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
                End If
                
            Case DIR_LEFT
                If MapNpc(MapNpcNum).num > 0 Then
                    MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - ((ElapsedTime / 1000) * (NPC(MapNpc(MapNpcNum).num).speed * SIZE_X))
                    If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
                End If
                
            Case DIR_RIGHT
                If MapNpc(MapNpcNum).num > 0 Then
                    MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + ((ElapsedTime / 1000) * (NPC(MapNpc(MapNpcNum).num).speed * SIZE_X))
                    If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
                End If
                
        End Select
    
        ' Check if completed walking over to the next tile
        If MapNpc(MapNpcNum).Moving > 0 Then
            If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
                If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If VXFRAME = False Then
                        If MapNpc(MapNpcNum).Step = 1 Then
                            MapNpc(MapNpcNum).Step = 3
                        Else
                            MapNpc(MapNpcNum).Step = 1
                        End If
                    Else
                        If MapNpc(MapNpcNum).Step = 0 Then
                            MapNpc(MapNpcNum).Step = 2
                        Else
                            MapNpc(MapNpcNum).Step = 0
                        End If
                    End If
                End If
            Else
                If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                    MapNpc(MapNpcNum).Moving = 0
                    If VXFRAME = False Then
                        If MapNpc(MapNpcNum).Step = 1 Then
                            MapNpc(MapNpcNum).Step = 3
                        Else
                            MapNpc(MapNpcNum).Step = 1
                        End If
                    Else
                        If MapNpc(MapNpcNum).Step = 0 Then
                            MapNpc(MapNpcNum).Step = 2
                        Else
                            MapNpc(MapNpcNum).Step = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ProcessNpcMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMapGetItem()
Dim buffer As New clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer

    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Player(MyIndex).MapGetTimer = GetTickCount
            buffer.WriteLong CMapGetItem
            SendData buffer.ToArray()
        End If
    End If

    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAttack()

Dim buffer As clsBuffer
Dim attackspeed As Long, X As Long, Y As Long, I As Long
Dim arma As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    If ControlDown Then
        Select Case Player(MyIndex).Dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
        
        'Interaccion Con Tile Cubos EaSee 0.4
                
                
        If GetTickCount > Player(MyIndex).EventTimer Then
            For I = 1 To Map.CurrentEvents
                If Map.MapEvents(I).Visible = 1 Then
                    If Map.MapEvents(I).X = X And Map.MapEvents(I).Y = Y Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CEvent
                        buffer.WriteLong I
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        Player(MyIndex).EventTimer = GetTickCount + 200
                    End If
                End If
            Next
        End If
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                  If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
                        ' projectile
                     If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Municion = 0 Then
                        Set buffer = New clsBuffer
                            buffer.WriteLong CProjecTileAttack
                            SendData buffer.ToArray()
                            Set buffer = Nothing
                            Exit Sub
                      Else
                        For I = 1 To MAX_INV
                             If GetPlayerInvItemNum(MyIndex, I) = Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Municion Then
                             'Call SetPlayerInvItemValue(MyIndex, I, GetPlayerInvItemValue(MyIndex, I) - 1)
                             Call SendUseItem(I)
                             Set buffer = New clsBuffer
                             buffer.WriteLong CProjecTileAttack
                             SendData buffer.ToArray()
                             Set buffer = Nothing
                             Exit Sub
                             End If
                       Next
                       AddText "No tienes Municiones", BrightRed
                    End If
                  End If
                 End If
                ' non projectile
                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
    End If
    
    
    
    
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsTryingToMove() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CanMove(Optional ByVal FollowingPlayer As Boolean, Optional ByVal Dir As Byte) As Boolean
Dim d As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If InEvent Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' make sure they're not in a shop
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' not in bank
    If InBank Then
        'CanMove = False
        'Exit Function
        InBank = False
        GUIWindow(GUI_BANK).Visible = False
    End If

    d = GetPlayerDir(MyIndex)

    If DirUp Or FollowingPlayer And Dir = 0 Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Up > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirDown Or FollowingPlayer And Dir = 1 Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Down > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirLeft Or FollowingPlayer And Dir = 2 Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Left > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    If DirRight Or FollowingPlayer And Dir = 3 Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If

        Else

            ' Check if they can warp to a new map
            If Map.Right > 0 Then
                Call MapEditorLeaveMap
                Call SendPlayerRequestNewMap
                GettingMap = True
                CanMoveNow = False
            End If

            CanMove = False
            Exit Function
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long
Dim Y As Long
Dim I As Long
Dim blockVar As Byte

    'QUICKCHANGE
    'On Error Resume Next
    'blockVar = Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CheckDirection = False
    
    ' check directional blocking
    If isDirBlocked(blockVar, Direction + 1) And Player(MyIndex).Walkthrough = False Then
        CheckDirection = True
        Exit Function
    End If

    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    ' Check to see if the map tile is blocked or not... agregamos Letreros EaSee 0.4 Version Cubos by ikicapitan
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Len(Map.Tile(X, Y).Mensaje) > 1 And Player(MyIndex).Walkthrough = False Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is a resource or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE And Player(MyIndex).Walkthrough = False Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the key door is open or not
    If Map.Tile(X, Y).Type = TILE_TYPE_KEY Then

        ' This actually checks if its open or not
        If TempTile(X, Y).DoorOpen = NO And Player(MyIndex).Walkthrough = False Then
            CheckDirection = True
            Exit Function
        End If
    End If
    
    
    
    ' Check to see if a npc is already on that tile
    For I = 1 To Npc_HighIndex
        If MapNpc(I).num > 0 And Player(MyIndex).Walkthrough = False Then
            If MapNpc(I).X = X Then
                If MapNpc(I).Y = Y Then
                    If Not Player(MyIndex).Walkthrough Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    
    ' Check for an event
    For I = 1 To Map.CurrentEvents
        If Map.MapEvents(I).Visible = 1 Then
            If Map.MapEvents(I).X = X Then
                If Map.MapEvents(I).Y = Y Then
                    If Map.MapEvents(I).Walkthrough = 0 And Player(MyIndex).Walkthrough = False Then
                        If Not Player(MyIndex).Walkthrough Then
                            CheckDirection = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    ' Don't worry about players if it's a safe zone
    If Map.Moral = MAP_MORAL_SAFE Then Exit Function

    ' Check to see if a player is already on that tile
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) And Player(MyIndex).Walkthrough = False Then
            If GetPlayerX(I) = X And Player(MyIndex).Walkthrough = False Then
                If GetPlayerY(I) = Y And Player(MyIndex).Walkthrough = False Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next I

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "checkDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub CheckMovement(Optional ByVal FollowingSomeone As Boolean = False, Optional ByVal Dir As Byte = 5)
Dim confusiondado As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ParalisisDuracion > 0 Then Exit Sub ' Paralisis EaSee 0.5
    If IsTryingToMove Or FollowingSomeone Then
    
        If CanMove(FollowingSomeone, Dir) Then

            ' Check if player has the shift key down for running
            If ShiftDown Then
                Player(MyIndex).Moving = MOVING_RUNNING
            Else
                Player(MyIndex).Moving = MOVING_WALKING
            End If

confusiondado = Rand(0, 1)
If ConfusionDuracion > 0 And confusiondado = 1 Then 'Confusion EaSee Engine 0.5

Select Case GetPlayerDir(MyIndex)
                
                Case DIR_UP
                    
    
                If Player(MyIndex).Y = Map.MaxY Then Exit Sub
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                Case DIR_DOWN
                
                If Player(MyIndex).Y = 0 Then Exit Sub
                      Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                Case DIR_LEFT
                    If Player(MyIndex).X = Map.MaxX Then Exit Sub
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                Case DIR_RIGHT

                    If Player(MyIndex).X = 0 Then Exit Sub
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
            End Select


Else


            Select Case GetPlayerDir(MyIndex)
                Case DIR_UP
    
                Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                Case DIR_DOWN
                    Call SendPlayerMove
                    Player(MyIndex).yOffset = PIC_Y * -1
                    Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                Case DIR_LEFT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                Case DIR_RIGHT
                    Call SendPlayerMove
                    Player(MyIndex).xOffset = PIC_X * -1
                    Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
            End Select

End If

            If Player(MyIndex).xOffset = 0 Then
                If Player(MyIndex).yOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Function isInBounds()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub UpdateDrawMapName()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    DrawMapNameX = GUIWindow(GUI_HOTBAR).X '((MAX_MAPX + 1) * PIC_X / 2) - (getWidth(Font_Default, Trim$(Map.name)) / 2)
    
    DrawMapNameY = GUIWindow(GUI_HOTBAR).Y + 40

    Select Case Map.Moral
        Case MAP_MORAL_NONE
            DrawMapNameColor = BrightRed
        Case MAP_MORAL_SAFE
            DrawMapNameColor = White
        Case Else
            DrawMapNameColor = White
    End Select

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then
        Exit Sub
    End If

    Call SendUseItem(InventoryItemSelected)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ForgetSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If SpellCD(spellSlot) > 0 Then
        AddText "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If SpellBuffer = spellSlot Then
        AddText "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong spellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "No spell here.", BrightRed
    End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal spellSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If spellSlot < 1 Or spellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    If SpellCD(spellSlot) > 0 Then
        AddText "Hechizo en enfriamiento!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(spellSlot) = 0 Then Exit Sub

    ' Check if player has enough MP
    If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(spellSlot)).MPCost Then
        Call AddText("MP insuficiente para lanzar " & Trim$(Spell(PlayerSpells(spellSlot)).name) & ".", BrightRed)
        Exit Sub
    End If

    If PlayerSpells(spellSlot) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If Player(MyIndex).Moving = 0 Then
                Set buffer = New clsBuffer
                buffer.WriteLong CCast
                buffer.WriteLong spellSlot
                SendData buffer.ToArray()
                Set buffer = Nothing
                SpellBuffer = spellSlot
                SpellBufferTimer = GetTickCount
            Else
                Call AddText("No puedes lanzar hechizos mientras caminas!", BrightRed)
            End If
        End If
    Else
        Call AddText("No hechizos aqui.", BrightRed)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ReDim TempTile(0 To Map.MaxX, 0 To Map.MaxY)

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearTempTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DevMsg(ByVal text As String, ByVal Color As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InGame Then
        If GetPlayerAccess(MyIndex) > ADMIN_DEVELOPER Then
            Call AddText(text, Color)
        End If
    End If

    Debug.Print text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DevMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function TwipsToPixels(ByVal twip_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = twip_val / Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function PixelsToTwips(ByVal pixel_val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = pixel_val * Screen.TwipsPerPixelY
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Int(Amount) < 10000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) <= 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) <= 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawPing()
Dim PingToDraw As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Syncing"
        Case 0 To 5
            PingToDraw = "Local"
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DrawPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal Color As Integer, ByVal MsgType As Byte, ByVal X As Long, ByVal Y As Long)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ActionMsgIndex = ActionMsgIndex + 1
    If ActionMsgIndex >= MAX_BYTE Then ActionMsgIndex = 1

    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .Color = Color
        .Type = MsgType
        .Created = GetTickCount
        .Scroll = 1
        .X = X
        .Y = Y
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Rand(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Rand(-8, 8)
    End If
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).Created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ActionMsg(Index).Message = vbNullString
    ActionMsg(Index).Created = 0
    ActionMsg(Index).Type = 0
    ActionMsg(Index).Color = 0
    ActionMsg(Index).Scroll = 0
    ActionMsg(Index).X = 0
    ActionMsg(Index).Y = 0
    
    ' find the new high index
    For I = MAX_BYTE To 1 Step -1
        If ActionMsg(I).Created > 0 Then
            Action_HighIndex = I + 1
            Exit For
        End If
    Next
    ' make sure we don't overflow
    If Action_HighIndex > MAX_BYTE Then Action_HighIndex = MAX_BYTE
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
Dim looptime As Long
Dim layer As Long
Dim FrameCount As Long
Dim lockindex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if doesn't exist then exit sub
    If AnimInstance(Index).Animation <= 0 Then Exit Sub
    If AnimInstance(Index).Animation >= MAX_ANIMATIONS Then Exit Sub
    
    For layer = 0 To 1
        If AnimInstance(Index).Used(layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(layer)
            
            ' if zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(layer) = 0 Then AnimInstance(Index).frameIndex(layer) = 1
            If AnimInstance(Index).LoopIndex(layer) = 0 Then AnimInstance(Index).LoopIndex(layer) = 1
            
            ' check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(layer) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimInstance(Index).frameIndex(layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(layer) = AnimInstance(Index).LoopIndex(layer) + 1
                    If AnimInstance(Index).LoopIndex(layer) > Animation(AnimInstance(Index).Animation).LoopCount(layer) Then
                        AnimInstance(Index).Used(layer) = False
                    Else
                        AnimInstance(Index).frameIndex(layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(layer) = AnimInstance(Index).frameIndex(layer) + 1
                End If
                AnimInstance(Index).Timer(layer) = GetTickCount
            End If
        End If
    Next
    
    ' if neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then ClearAnimInstance (Index)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenShop(ByVal shopnum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InShop = shopnum
    ShopAction = 0
    GUIWindow(GUI_SHOP).Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemNum(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If bankslot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If bankslot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(bankslot).num
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemNum(ByVal bankslot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(bankslot).num = itemNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetBankItemValue(ByVal bankslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetBankItemValue = Bank.Item(bankslot).Value
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetBankItemValue(ByVal bankslot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(bankslot).Value = ItemValue
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' BitWise Operators for directional blocking
Public Sub setDirBlock(ByRef blockVar As Byte, ByRef Dir As Byte, ByVal block As Boolean)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If block Then
        blockVar = blockVar Or (2 ^ Dir)
    Else
        blockVar = blockVar And Not (2 ^ Dir)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function isDirBlocked(ByRef blockVar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not blockVar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Sub PlayMapSound(ByVal X As Long, ByVal Y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim soundName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If entityNum <= 0 Then Exit Sub
    
    ' find the sound
    Select Case entityType
        ' animations
        Case SoundEntity.seAnimation
            If entityNum > MAX_ANIMATIONS Then Exit Sub
            soundName = Trim$(Animation(entityNum).sound)
            
        ' items
        Case SoundEntity.seItem
            If entityNum > MAX_ITEMS Then Exit Sub
            soundName = Trim$(Item(entityNum).sound)
        ' npcs
        Case SoundEntity.seNpc
            If entityNum > MAX_NPCS Then Exit Sub
            soundName = Trim$(NPC(entityNum).sound)
        ' resources
        Case SoundEntity.seResource
            If entityNum > MAX_RESOURCES Then Exit Sub
            soundName = Trim$(Resource(entityNum).sound)
        ' spells
        Case SoundEntity.seSpell
            If entityNum > MAX_SPELLS Then Exit Sub
            soundName = Trim$(Spell(entityNum).sound)
        ' other
        Case Else
            Exit Sub
    End Select
    
    ' exit out if it's not set
    If Trim$(soundName) = "Ninguno." Then Exit Sub

    ' play the sound
    PlaySound soundName, X, Y
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PlayMapSound", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Dialogue(ByVal diTitle As String, ByVal diText As String, ByVal diIndex As Long, Optional ByVal isYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' exit out if we've already got a dialogue open
    If dialogueIndex > 0 Then Exit Sub
    
    ' set global dialogue index
    dialogueIndex = diIndex
    
    ' set the global dialogue data
    dialogueData1 = Data1

    ' set the captions
    Dialogue_TitleCaption = diTitle
    Dialogue_TextCaption = diText
    
    ' show/hide buttons
    If Not isYesNo Then
        Dialogue_ButtonVisible(1) = False ' Yes button
        Dialogue_ButtonVisible(2) = True ' Okay button
        Dialogue_ButtonVisible(3) = False ' No button
    Else
        Dialogue_ButtonVisible(1) = True ' Yes button
        Dialogue_ButtonVisible(2) = False ' Okay button
        Dialogue_ButtonVisible(3) = True ' No button
    End If
    
    ' show the dialogue box
    GUIWindow(GUI_DIALOGUE).Visible = True
    inChat = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub dialogueHandler(ByVal Index As Long)
    ' find out which button
    If Index = 1 Then ' okay button
        ' dialogue index
        Select Case dialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' yes button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                ForgetSpell dialogueData1
            Case DIALOGUE_TYPE_PARTY
                SendAcceptParty
        End Select
    ElseIf Index = 3 Then ' no button
        ' dialogue index
        Select Case dialogueIndex
            Case DIALOGUE_TYPE_TRADE
                SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                SendDeclineParty
        End Select
    End If
End Sub

Sub ProcessEventMovement(ByVal ID As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if NPC is walking, and if so process moving them over
    If Map.MapEvents(ID).Moving = 1 Then
        
        Select Case Map.MapEvents(ID).Dir
            Case DIR_UP
                Map.MapEvents(ID).yOffset = Map.MapEvents(ID).yOffset - ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).yOffset < 0 Then Map.MapEvents(ID).yOffset = 0
                
            Case DIR_DOWN
                Map.MapEvents(ID).yOffset = Map.MapEvents(ID).yOffset + ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).yOffset > 0 Then Map.MapEvents(ID).yOffset = 0
                
            Case DIR_LEFT
                Map.MapEvents(ID).xOffset = Map.MapEvents(ID).xOffset - ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).xOffset < 0 Then Map.MapEvents(ID).xOffset = 0
                
            Case DIR_RIGHT
                Map.MapEvents(ID).xOffset = Map.MapEvents(ID).xOffset + ((ElapsedTime / 1000) * (Map.MapEvents(ID).MovementSpeed * SIZE_X))
                If Map.MapEvents(ID).xOffset > 0 Then Map.MapEvents(ID).xOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If Map.MapEvents(ID).Moving > 0 Then
            If Map.MapEvents(ID).Dir = DIR_RIGHT Or Map.MapEvents(ID).Dir = DIR_DOWN Then
                If (Map.MapEvents(ID).xOffset >= 0) And (Map.MapEvents(ID).yOffset >= 0) Then
                    Map.MapEvents(ID).Moving = 0
                    If VXFRAME = False Then
                        If Map.MapEvents(ID).Step = 1 Then
                            Map.MapEvents(ID).Step = 3
                        Else
                            Map.MapEvents(ID).Step = 1
                        End If
                    Else
                        If Map.MapEvents(ID).Step = 0 Then
                            Map.MapEvents(ID).Step = 2
                        Else
                            Map.MapEvents(ID).Step = 0
                        End If
                    End If
                End If
            Else
                If (Map.MapEvents(ID).xOffset <= 0) And (Map.MapEvents(ID).yOffset <= 0) Then
                    Map.MapEvents(ID).Moving = 0
                    If VXFRAME = False Then
                        If Map.MapEvents(ID).Step = 1 Then
                            Map.MapEvents(ID).Step = 3
                        Else
                            Map.MapEvents(ID).Step = 1
                        End If
                    Else
                        If Map.MapEvents(ID).Step = 0 Then
                            Map.MapEvents(ID).Step = 2
                        Else
                            Map.MapEvents(ID).Step = 0
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ProcessEventMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetColorString(Color As Long)
    Select Case Color
        Case 0
            GetColorString = "Black"
        Case 1
            GetColorString = "Blue"
        Case 2
            GetColorString = "Green"
        Case 3
            GetColorString = "Cyan"
        Case 4
            GetColorString = "Red"
        Case 5
            GetColorString = "Magenta"
        Case 6
            GetColorString = "Brown"
        Case 7
            GetColorString = "Grey"
        Case 8
            GetColorString = "Dark Grey"
        Case 9
            GetColorString = "Bright Blue"
        Case 10
            GetColorString = "Bright Green"
        Case 11
            GetColorString = "Bright Cyan"
        Case 12
            GetColorString = "Bright Red"
        Case 13
            GetColorString = "Pink"
        Case 14
            GetColorString = "Yellow"
        Case 15
            GetColorString = "White"

    End Select
End Function

Sub ClearEventChat()
Dim I As Long
    For I = 1 To 4
        chatOpt(I) = vbNullString
    Next
    chatText = vbNullString
    GUIWindow(GUI_EVENTCHAT).Visible = False
    inChat = False
End Sub

Public Sub MenuLoop()
    ' If debug mode, handle error then exit out
    On Error GoTo ErrorHandler
restartmenuloop:
    ' *** Start GameLoop ***
    CountMenu = 0
    PosMenu = 10
    Alphamenu = 0
    MoveMenu = 1
    VisibleTextMenu = 0
    Do While Not InGame
        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call DrawGDI
        DoEvents
    Loop

    ' Error handler
    Exit Sub
ErrorHandler:
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost
        GoTo restartmenuloop
    ElseIf Options.Debug = 1 Then
        HandleError "MenuLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
        Err.Clear
        Exit Sub
    End If
End Sub

Sub ProcessWeather()
Dim I As Long
    If CurrentWeather > 0 Then
        I = Rand(1, 101 - CurrentWeatherIntensity)
        If I = 1 Then
            'Add a new particle
            For I = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(I).InUse = False Then
                    If Rand(1, 2) = 1 Then
                        WeatherParticle(I).InUse = True
                        WeatherParticle(I).Type = CurrentWeather
                        WeatherParticle(I).Velocity = Rand(8, 14)
                        WeatherParticle(I).X = (TileView.Left * 32) - 32
                        WeatherParticle(I).Y = (TileView.Top * 32) + Rand(-32, frmMain.ScaleHeight)
                    Else
                        WeatherParticle(I).InUse = True
                        WeatherParticle(I).Type = CurrentWeather
                        WeatherParticle(I).Velocity = Rand(10, 15)
                        WeatherParticle(I).X = (TileView.Left * 32) + Rand(-32, frmMain.ScaleWidth)
                        WeatherParticle(I).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WEATHER_TYPE_STORM Then
        I = Rand(1, 400 - CurrentWeatherIntensity)
        If I = 1 Then
            'Draw Thunder
            DrawThunder = Rand(15, 22)
            PlaySound Sound_Thunder, -1, -1
        End If
    End If
    
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).X > TileView.Right * 32 Or WeatherParticle(I).Y > TileView.Bottom * 32 Then
                WeatherParticle(I).InUse = False
            Else
                WeatherParticle(I).X = WeatherParticle(I).X + WeatherParticle(I).Velocity
                WeatherParticle(I).Y = WeatherParticle(I).Y + WeatherParticle(I).Velocity
            End If
        End If
    Next
End Sub

Public Sub AddChatBubble(ByVal target As Long, ByVal targetType As Byte, ByVal Msg As String, ByVal colour As Long)
Dim I As Long, Index As Long

    ' set the global index
    chatBubbleIndex = chatBubbleIndex + 1
    If chatBubbleIndex < 1 Or chatBubbleIndex > MAX_BYTE Then chatBubbleIndex = 1
    
    ' default to new bubble
    Index = chatBubbleIndex
    
    ' loop through and see if that player/npc already has a chat bubble
    For I = 1 To MAX_BYTE
        If chatBubble(I).targetType = targetType Then
            If chatBubble(I).target = target Then
                ' reset master index
                If chatBubbleIndex > 1 Then chatBubbleIndex = chatBubbleIndex - 1
                ' we use this one now, yes?
                Index = I
                Exit For
            End If
        End If
    Next
    
    ' set the bubble up
    With chatBubble(Index)
        .target = target
        .targetType = targetType
        .Msg = Msg
        .colour = colour
        .Timer = GetTickCount
        .active = True
    End With
End Sub

Public Sub ScrollEditor()
    Scroll_Draw = True
    frmMain.tmrScrollEditor.Enabled = False
    Scroll_Timer = 0
    frmMain.tmrScrollEditor.Enabled = True
End Sub

Public Function IsBankItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsBankItem = 0
    
    For I = 1 To MAX_BANK
        If Not emptySlot Then
            If GetBankItemNum(I) <= 0 And GetBankItemNum(I) > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_BANK).Y + BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_BANK).X + BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    
                    IsBankItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsBankItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
Dim I As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            Top = GUIWindow(GUI_SHOP).Y + ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
            Left = GUIWindow(GUI_SHOP).X + ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))

            If X >= Left And X <= Left + 32 Then
                If Y >= Top And Y <= Top + 32 Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsShopItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim tempRec As RECT
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsEqItem = 0

    For I = 1 To Equipment.Equipment_Count - 1

        If GetPlayerEquipment(MyIndex, I) > 0 And GetPlayerEquipment(MyIndex, I) <= MAX_ITEMS Then

            With tempRec
            
             Select Case Item(GetPlayerEquipment(MyIndex, I)).Type
        
        Case 1 'Arma
                
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 * 2 + 6
        .Left = GUIWindow(GUI_CHARACTER).X + 57
        
        Case 2 'Armadura JapiDura
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 + 3
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 * 4 + 13
        
        Case 3 'Casco corte Cartilos Bala
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 * 4 + 13
        
        Case 4 'Piernas (no te pueden faltar man)
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 * 2 + 6
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 * 4 + 13
        
        
        Case 5 'Botas porque hay sufragio universal
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 * 3 + 8
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 * 2 + 4
        
        Case 6 'Guantes no te aguantes
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 * 3 + 8
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 * 3 + 4
        
        Case 7 'Anillo
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 * 3 + 8
        .Left = GUIWindow(GUI_CHARACTER).X + 57 + 32 + 4
        
        Case 8 'Collar
        .Top = GUIWindow(GUI_CHARACTER).Y + 43
        .Left = GUIWindow(GUI_CHARACTER).X + 57
        
        Case 9 'Escudo
        
        .Top = GUIWindow(GUI_CHARACTER).Y + 43 + 32 + 6
        .Left = GUIWindow(GUI_CHARACTER).X + 57
        
        End Select

            
            
        .Bottom = .Top + PIC_Y
        .Right = .Left + PIC_X
            End With

            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsEqItem = I
                    Exit Function
                End If
            End If
        End If

    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsEqItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsInvItem(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsInvItem = 0

    For I = 1 To MAX_INV
        
        If Not emptySlot Then
            If GetPlayerInvItemNum(MyIndex, I) <= 0 Or GetPlayerInvItemNum(MyIndex, I) > MAX_ITEMS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_INVENTORY).Y + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_INVENTORY).X + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Left = .Left + 12 ' Fix al offset
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsInvItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsInvItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single, Optional ByVal emptySlot As Boolean = False) As Long
Dim tempRec As RECT, skipThis As Boolean
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsPlayerSpell = 0

    For I = 1 To MAX_PLAYER_SPELLS

        If Not emptySlot Then
            If PlayerSpells(I) <= 0 And PlayerSpells(I) > MAX_PLAYER_SPELLS Then skipThis = True
        End If

        If Not skipThis Then
            With tempRec
                .Top = GUIWindow(GUI_SPELLS).Y + SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_SPELLS).X + SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsPlayerSpell = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsPlayerSpell", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean, Optional ByVal emptySlot As Boolean = False) As Long
    Dim tempRec As RECT, skipThis As Boolean
    Dim I As Long
    Dim IsTradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsTradeItem = 0

    For I = 1 To MAX_INV
    
        If Yours Then
            IsTradeNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
        Else
            IsTradeNum = TradeTheirOffer(I).num
        End If
        
        If Not emptySlot Then
            If IsTradeNum <= 0 Or IsTradeNum > MAX_ITEMS Then skipThis = True
        End If
        
        If Not skipThis Then
             With tempRec
                .Top = GUIWindow(GUI_TRADE).Y + 31 + InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = GUIWindow(GUI_TRADE).X + 29 + InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With
    
            If X >= tempRec.Left And X <= tempRec.Right Then
                If Y >= tempRec.Top And Y <= tempRec.Bottom Then
                    IsTradeItem = I
                    Exit Function
                End If
            End If
        End If
        skipThis = False
    Next

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsTradeItem", "frmGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single) As Long
Dim Top As Long, Left As Long
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    IsHotbarSlot = 0

    For I = 1 To MAX_HOTBAR
        Top = GUIWindow(GUI_HOTBAR).Y + HotbarTop
        Left = GUIWindow(GUI_HOTBAR).X + HotbarLeft + ((HotbarOffsetX + 32) * (((I - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                IsHotbarSlot = I
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    CensorWord = String(Len(sString), "*")
End Function
Public Sub setOptionsState()
    ' music
    If Options.Music = 1 Then
        Buttons(26).state = 2
        Buttons(27).state = 0
    Else
        Buttons(26).state = 0
        Buttons(27).state = 2
    End If
    
    ' sound
    If Options.sound = 1 Then
        Buttons(28).state = 2
        Buttons(29).state = 0
    Else
        Buttons(28).state = 0
        Buttons(29).state = 2
    End If
    
    ' debug
    If Options.Debug = 1 Then
        Buttons(30).state = 2
        Buttons(31).state = 0
    Else
        Buttons(30).state = 0
        Buttons(31).state = 2
    End If
    
    ' levels
    If Options.Lvls = 1 Then
        Buttons(32).state = 2
        Buttons(33).state = 0
    Else
        Buttons(32).state = 0
        Buttons(33).state = 2
    End If
    
    ' minimap
    If Options.MiniMap = 1 Then
        Buttons(59).state = 2
        Buttons(60).state = 0
    Else
        Buttons(59).state = 0
        Buttons(60).state = 2
    End If
    
    ' buttons
    If Options.Buttons = 1 Then
        Buttons(61).state = 2
        Buttons(62).state = 0
    Else
        Buttons(61).state = 0
        Buttons(62).state = 2
    End If
End Sub

Public Sub ScrollChatBox(ByVal Direction As Byte)
    ' do a quick exit if we don't have enough text to scroll
    If totalChatLines < 8 Then
        ChatScroll = 8
        UpdateChatArray
        Exit Sub
    End If
    ' actually scroll
    If Direction = 0 Then ' up
        ChatScroll = ChatScroll + 1
    Else ' down
        ChatScroll = ChatScroll - 1
    End If
    ' scrolling down
    If ChatScroll < 8 Then ChatScroll = 8
    ' scrolling up
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    ' update the array
    UpdateChatArray
End Sub

Public Sub CheckAttackCubo() 'EaSee 0.9

Dim buffer As clsBuffer
Dim attackspeed As Long, X As Long, Y As Long, I As Long
Dim arma As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    If CanMoveNow Then
    
        Select Case Player(MyIndex).Dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
        
        'Interaccion Con Tile Cubos EaSee 0.4
                
                
        If GetTickCount > Player(MyIndex).EventTimer Then
            For I = 1 To Map.CurrentEvents
                If Map.MapEvents(I).Visible = 1 Then
                    If Map.MapEvents(I).X = X And Map.MapEvents(I).Y = Y Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CEvent
                        buffer.WriteLong I
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        Player(MyIndex).EventTimer = GetTickCount + 200
                    End If
                End If
            Next
        End If
    
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack

        ' speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(MyIndex, Weapon)).speed
        Else
            attackspeed = 1000
        End If

        If Player(MyIndex).AttackTimer + attackspeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then

                With Player(MyIndex)
                    .Attacking = 1
                    .AttackTimer = GetTickCount
                End With

                If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
                  If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Pic > 0 Then
                        ' projectile
                     If Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Municion = 0 Then
                        Set buffer = New clsBuffer
                            buffer.WriteLong CProjecTileAttack
                            SendData buffer.ToArray()
                            Set buffer = Nothing
                            Exit Sub
                      Else
                        For I = 1 To MAX_INV
                             If GetPlayerInvItemNum(MyIndex, I) = Item(GetPlayerEquipment(MyIndex, Weapon)).ProjecTile.Municion Then
                             'Call SetPlayerInvItemValue(MyIndex, I, GetPlayerInvItemValue(MyIndex, I) - 1)
                             Call SendUseItem(I)
                             Set buffer = New clsBuffer
                             buffer.WriteLong CProjecTileAttack
                             SendData buffer.ToArray()
                             Set buffer = Nothing
                             Exit Sub
                             End If
                       Next
                       AddText "No tienes Municiones", BrightRed
                    End If
                  End If
                 End If
                ' non projectile
                Set buffer = New clsBuffer
                buffer.WriteLong CAttackCubo
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
    End If
    
    
    
    
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
