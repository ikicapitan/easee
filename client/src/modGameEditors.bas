Attribute VB_Name = "modGameEditors"
Option Explicit
Public cpEvent As EventRec
Const LB_SETHORIZONTALEXTENT = &H194
Private Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public EventList() As EventListRec

' ////////////////
' // Map Editor //
' ////////////////
Public Sub MapEditorInit()
Dim I As Long
Dim smusic() As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' set the width
    frmEditor_Map.Width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.Visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.max = NumTileSets
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "Ninguno"
    For I = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem I & ": " & Shop(I).name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorProperties()
Dim X As Long
Dim Y As Long
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "Ninguno."
    For I = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(I)
    Next
    frmEditor_MapProperties.cmbSound.Clear
    frmEditor_MapProperties.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_MapProperties.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .txtName.text = Trim$(Map.name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For I = 0 To .lstMusic.ListCount
                If .lstMusic.List(I) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = I
                End If
            Next
        End If
        
        If .cmbSound.ListCount >= 0 Then
            .cmbSound.ListIndex = 0
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Map.BGS) Then
                    .cmbSound.ListIndex = I
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.text = CStr(Map.Up)
        .txtDown.text = CStr(Map.Down)
        .txtLeft.text = CStr(Map.Left)
        .txtRight.text = CStr(Map.Right)
        .cmbMoral.ListIndex = Map.Moral
        .txtBootMap.text = CStr(Map.BootMap)
        .txtBootX.text = CStr(Map.BootX)
        .txtBootY.text = CStr(Map.BootY)
        
        .CmbWeather.ListIndex = Map.Weather
        .scrlWeatherIntensity.Value = Map.WeatherIntensity
        
        .ScrlFog.Value = Map.Fog
        .ScrlFogSpeed.Value = Map.FogSpeed
        .scrlFogOpacity.Value = Map.FogOpacity
        
        .ScrlR.Value = Map.Red
        .ScrlG.Value = Map.Green
        .ScrlB.Value = Map.Blue
        .scrlA.Value = Map.alpha

        ' show the map npcs
        .lstNpcs.Clear
        For X = 1 To MAX_MAP_NPCS
            If Map.NPC(X) > 0 Then
            .lstNpcs.AddItem X & ": " & Trim$(NPC(Map.NPC(X)).name)
            Else
                .lstNpcs.AddItem X & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For X = 1 To MAX_NPCS
            .cmbNpc.AddItem X & ": " & Trim$(NPC(X).name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = Map.NPC(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Mapa Actual: " & GetPlayerMap(MyIndex)
        .txtMaxX.text = Map.MaxX
        .txtMaxY.text = Map.MaxY
        .chkDrop.Value = Map.DropItemsOnDeath
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorProperties", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If theAutotile > 0 Then
        With Map.Tile(X, Y)
            ' set layer
            .layer(CurLayer).X = EditorTileX
            .layer(CurLayer).Y = EditorTileY
            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = theAutotile
            CacheRenderState X, Y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With Map.Tile(X, Y)
            ' set layer
            .layer(CurLayer).X = EditorTileX
            .layer(CurLayer).Y = EditorTileY
            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = 0
            CacheRenderState X, Y, CurLayer
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For Y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .layer(CurLayer).X = EditorTileX + X2
                            .layer(CurLayer).Y = EditorTileY + Y2
                            .layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                            .Autotile(CurLayer) = 0
                            CacheRenderState X, Y, CurLayer
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorSetTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal movedMouse As Boolean = True)
Dim I As Long
Dim CurLayer As Long
Dim tmpDir As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.optEvent.Value Then
            If frmEditor_Events.Visible = False Then
                AddEvent CurX, CurY
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then .Type = TILE_TYPE_BLOCKED
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key
                If frmEditor_Map.optKey.Value Then
                    .Type = TILE_TYPE_KEY
                    .Data1 = KeyEditorNum
                    .Data2 = KeyEditorTake
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' key open
                If frmEditor_Map.optKeyOpen.Value Then
                    .Type = TILE_TYPE_KEYOPEN
                    .Data1 = KeyOpenEditorX
                    .Data2 = KeyOpenEditorY
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' door
                If frmEditor_Map.optDoor.Value Then
                    .Type = TILE_TYPE_DOOR
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = ""
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = ""
                End If
                ' sound
                If frmEditor_Map.optSound.Value Then
                    .Type = TILE_TYPE_SOUND
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = MapEditorSound
                End If
                ' player spawn
                If frmEditor_Map.optPlayerSpawn.Value Then .Type = TILE_TYPE_PLAYERSPAWN
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            X = X - ((X \ 32) * 32)
            Y = Y - ((Y \ 32) * 32)
            ' see if it hits an arrow
            For I = 1 To 4
                If X >= DirArrowX(I) And X <= DirArrowX(I) + 8 Then
                    If Y >= DirArrowY(I) And Y <= DirArrowY(I) + 8 Then
                        ' flip the value.
                        setDirBlock Map.Tile(CurX, CurY).DirBlock, CByte(I), Not isDirBlocked(Map.Tile(CurX, CurY).DirBlock, CByte(I))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With Map.Tile(CurX, CurY)
                ' clear layer
                .layer(CurLayer).X = 0
                .layer(CurLayer).Y = 0
                .layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                CacheRenderState X, Y, CurLayer
            End With
        ElseIf frmEditor_Map.optEvent.Value Then
            Call DeleteEvent(CurX, CurY)
        ElseIf frmEditor_Map.optAttribs.Value Then
            With Map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With

        End If
    End If

    CacheResources
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ PIC_X
        EditorTileY = Y \ PIC_Y
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorChooseTile", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        X = (X \ PIC_X) + 1
        Y = (Y \ PIC_Y) + 1
        ' check it's not out of bounds
        If X < 0 Then X = 0
        If X > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X Then X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width / PIC_X
        If Y < 0 Then Y = 0
        If Y > Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y Then Y = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height / PIC_Y
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
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorDrag", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' horizontal scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width < frmEditor_Map.picBack.Width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
    End If
    
    ' vertical scrolling
    If Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height < frmEditor_Map.picBack.Height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorTileScroll", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call SendMap
    SendMapRespawn
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorSend", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong 1
    SendData buffer.ToArray()
    SendMapRespawn
    InMapEditor = False
    Unload frmEditor_Map
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim I As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If Msgbox("Deseas eliminar esta capa?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).layer(CurLayer).X = 0
                Map.Tile(X, Y).layer(CurLayer).Y = 0
                Map.Tile(X, Y).layer(CurLayer).Tileset = 0
                CacheRenderState X, Y, CurLayer
            Next
        Next
        
        initAutotiles
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorClearLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim I As Long
Dim X As Long
Dim Y As Long
Dim CurLayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If Msgbox("Deseas con tu alma rellenar esta capa?", vbYesNo, Options.Game_Name) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).layer(CurLayer).X = EditorTileX
                Map.Tile(X, Y).layer(CurLayer).Y = EditorTileY
                Map.Tile(X, Y).layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                Map.Tile(X, Y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                CacheRenderState X, Y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
    

    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorFillLayer", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearAttribs()
Dim X As Long
Dim Y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Msgbox("Vaciar los atributos del mapa?", vbYesNo, Options.Game_Name) = vbYes Then

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                Map.Tile(X, Y).Type = 0
            Next
        Next

    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorClearAttribs", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorLeaveMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If InMapEditor Then
        If Msgbox("Guardar Cambios al Mapa Actual?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorLeaveMap", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorPlaceRandomTile(ByVal X As Long, Y As Long)
Dim I As Long
Dim CurLayer As Long

' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' find which layer we're on
    For I = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(I).Value Then
            CurLayer = I
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub

    If frmEditor_Map.optLayers.Value Then
        If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
            MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
        Else ' multi tile!
            If frmEditor_Map.scrlAutotile.Value = 0 Then
                MapEditorSetTile X, Y, CurLayer, True
            Else
                MapEditorSetTile X, Y, CurLayer, , frmEditor_Map.scrlAutotile.Value
            End If
        End If
    End If

    CacheResources

' Error handler
Exit Sub
ErrorHandler:
    HandleError "MapEditorMouseDown", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim I As Long
Dim SoundSet As Boolean
Dim sText As String
Dim F_CHAR As String * 1

    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Item.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    frmEditor_Item.scrlNum.max = NumTileSets
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(I)
    Next
    
    ' finished populating

    With Item(EditorIndex)
    
        frmEditor_Item.txtName.text = Trim$(.name)
        If .Pic > frmEditor_Item.scrlPic.max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.text = Trim$(.Desc)
        frmEditor_Item.chkStackable.Value = .Stackable
        frmEditor_Item.chkHanded.Value = .Handed
        frmEditor_Item.cmdCombatType.ListIndex = .CombatTypeReq
        frmEditor_Item.lblCombatType.Caption = "Tipo: "
        frmEditor_Item.scrlCombatLvl.max = MAX_COMBAT_LEVEL
        frmEditor_Item.scrlCombatLvl = .CombatLvlReq
        frmEditor_Item.lblCombatLvl.Caption = "Nivel de Combate: " & .CombatLvlReq
        frmEditor_Item.cmbSkill.ListIndex = .SkillReq - 1
        
        'asigna automaticamente 32x64 si esta vacio
        If Item(EditorIndex).Cubo32 = False And Item(EditorIndex).Cubo64 = False Then
        frmEditor_Item.opt32x64.Value = True
        Item(EditorIndex).Cubo64 = True
        End If
                
        sText = Replace$(.Book.text, F_CHAR, vbNullString)
        If Len(sText) > 0 Then
            frmEditor_Item.rtbBookText.text = Trim$(.Book.text)
        ElseIf Len(sText) = 0 Then
            frmEditor_Item.rtbBookText.text = Trim$("Página 1")
        End If
        
        sText = Replace$(.Book.Text2, F_CHAR, vbNullString)
        If Len(sText) > 0 Then
            frmEditor_Item.rtbBookText2.text = Trim$(.Book.Text2)
        ElseIf Len(sText) = 0 Then
            frmEditor_Item.rtbBookText2.text = Trim$("Página 2")
        End If
        
        If .Type = ITEM_TYPE_WEAPON Or .Type = ITEM_TYPE_SPELL Or .Type = ITEM_TYPE_PICACUBOS Then
            frmEditor_Item.cmdCombatType.Enabled = True
            frmEditor_Item.scrlCombatLvl.Enabled = True
        Else
            frmEditor_Item.cmdCombatType.Enabled = False
            frmEditor_Item.scrlCombatLvl.Enabled = False
        End If
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For I = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(I) = Trim$(.sound) Then
                    frmEditor_Item.cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PICACUBOS) Then
            frmEditor_Item.fraEquipment.Visible = True
            frmEditor_Item.scrlDamage.Value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3

            If .speed < 100 Then .speed = 100
            frmEditor_Item.scrlSpeed.Value = .speed
            
            ' loop for stats
            For I = 1 To Stats.Stat_Count - 2
                frmEditor_Item.scrlStatBonus(I).Value = .Add_Stat(I)
            Next
            
            For I = 6 To 11
                Select Case I
                    Case 6
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Light_Dmg
                    Case 7
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Light_Res
                    Case 8
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Dark_Res
                    Case 9
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Dark_Dmg
                    Case 10
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Neut_Dmg
                    Case 11
                        frmEditor_Item.scrlStatBonus(I).Value = .Element_Neut_Res
                End Select
            Next I
            
            frmEditor_Item.scrlPaperdoll = .Paperdoll
            
            If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON Or frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PICACUBOS Then
                frmEditor_Item.Frame4.Visible = True
                With Item(EditorIndex).ProjecTile
                      frmEditor_Item.scrlProjectileDamage.Value = .Damage
                      frmEditor_Item.scrlProjectilePic.Value = .Pic
                      frmEditor_Item.scrlProjectileRange.Value = .Range
                      frmEditor_Item.scrlProjectileSpeed.Value = .speed
                End With
            End If
        Else
            frmEditor_Item.fraEquipment.Visible = False
            frmEditor_Item.Frame4.Visible = False
        End If

        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.Visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
            frmEditor_Item.scrlCastSpell.Value = .CastSpell
            frmEditor_Item.chkInstant.Value = .instaCast
        Else
            frmEditor_Item.fraVitals.Visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.Visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.Visible = False
        End If
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_MUNICION) Then
        End If
        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CUBO Then
            'datos de cubo a cargar y enviar del cubo (sino nunca se guardan los cambios papi)
            If Item(EditorIndex).CuboTileN = 0 Then Item(EditorIndex).CuboTileN = 1 'El control le puse minimo 1 por lo que da error si es 0
            If Item(EditorIndex).CuboCapa1 = 0 Then Item(EditorIndex).CuboCapa1 = 2 'Cuando creamos un nuevo item es 0, devolveria siempre error
            If Item(EditorIndex).CuboCapa2 = 0 Then Item(EditorIndex).CuboCapa2 = 2
            If Item(EditorIndex).CuboSFX1 = 0 Then Item(EditorIndex).CuboSFX1 = 1
            If Item(EditorIndex).CuboSFX2 = 0 Then Item(EditorIndex).CuboSFX2 = 1
            If Item(EditorIndex).CuboAnimacion = 0 Then Item(EditorIndex).CuboAnimacion = 1
            If Item(EditorIndex).CuboObjeto = 0 Then Item(EditorIndex).CuboObjeto = 1
                        
            frmEditor_Item.scrlNum.Value = .CuboTileN 'Atributos de cada cubo, datos almacenados en Item
            frmEditor_Item.scrlX.Value = .CuboTileX
            frmEditor_Item.scrlY.Value = .CuboTileY
            frmEditor_Item.opt32x32.Value = .Cubo32
            frmEditor_Item.opt32x64.Value = .Cubo64
            frmEditor_Item.scrlCuboCapa.Value = .CuboCapa1
            frmEditor_Item.scrlCuboCapa2.Value = .CuboCapa2
            frmEditor_Item.scrlCuboSupTipo.Value = .CuboSupTipo
            frmEditor_Item.scrlCuboInfTipo.Value = .CuboInfTipo
            frmEditor_Item.txtCuboMapa.text = .CuboMapa
            frmEditor_Item.txtCuboX.text = .CuboMapaX
            frmEditor_Item.txtCuboY.text = .CuboMapaY
            frmEditor_Item.txtCuboGolpe.text = .CuboGolpe
            frmEditor_Item.txtCuboDureza.text = .CuboDureza
            frmEditor_Item.scrllAnim.Value = .CuboAnimacion
            frmEditor_Item.scrllSFX1.Value = .CuboSFX1
            frmEditor_Item.scrllSFX2.Value = .CuboSFX2
            frmEditor_Item.scrllDropear.Value = .CuboObjeto
            'Fin de carga-almacenamiento de datos
            
            Else
            
        End If

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' loop for stats
        For I = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(I).Value = .Stat_Req(I)
        Next
        
        ' Build cmbClassReq
        frmEditor_Item.cmbClassReq.Clear
        frmEditor_Item.cmbClassReq.AddItem "Ninguno"

        For I = 1 To Max_Classes
            frmEditor_Item.cmbClassReq.AddItem Class(I).name
        Next

        frmEditor_Item.cmbClassReq.ListIndex = .ClassReq
        ' Info
        frmEditor_Item.scrlPrice.Value = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With
    
    Item_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ItemEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ComboEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Combos.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Combos.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    ' not sure if I need this in here or not... It's staying though :p
    If Not hasPopulated Then
        PopulateLists
    End If
    
    ' add the array to the combo
    With frmEditor_Combos
        If Combo(EditorIndex).Item_1 > 0 Then
            .scrlItem1.Value = Combo(EditorIndex).Item_1
        Else
            .scrlItem1.Value = 1
        End If
        If Combo(EditorIndex).Item_2 > 0 Then
            .scrlItem2.Value = Combo(EditorIndex).Item_2
        Else
            .scrlItem2.Value = 1
        End If
        .cmbSkill.ListIndex = Combo(EditorIndex).Skill
        .scrlSkillLevel.Value = Combo(EditorIndex).SkillLevel
        .scrlLevel.Value = Combo(EditorIndex).Level
        If Combo(EditorIndex).Item_Given(.scrlIndex.Value) > 0 Then
            .scrlGive.Value = Combo(EditorIndex).Item_Given(.scrlIndex.Value)
        Else
            .scrlGive.Value = 1
        End If
        .scrlGiveVal = Combo(EditorIndex).Item_Given_Val(.scrlIndex.Value)
        .scrlSkillExp.Value = Combo(EditorIndex).GiveSkill_Exp
        .chkItem1.Value = Combo(EditorIndex).Take_Item1
        .chkItem2.Value = Combo(EditorIndex).Take_Item2
        .cmbGSkill.ListIndex = Combo(EditorIndex).GiveSkill
        .cmbItems1.ListIndex = Combo(EditorIndex).ReqItem1
        .cmbItems2.ListIndex = Combo(EditorIndex).ReqItem2
        .scrlItemVal1.Value = Combo(EditorIndex).ReqItemVal1
        .scrlItemVal2.Value = Combo(EditorIndex).ReqItemVal2
        .chkTake1.Value = Combo(EditorIndex).Take_ReqItem1
        .chkTake2.Value = Combo(EditorIndex).Take_ReqItem2
        COMBO_EDITOR_ITEM_INDEX = .scrlIndex.Value
    End With
    
    Combo_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ComboEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ITEMS
        If Item_Changed(I) Then
            Call SendSaveItem(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Item
    Editor = 0
    ClearChanged_Item
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ItemEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ComboEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_COMBO
        If Combo_Changed(I) Then
            Call SendSaveCombo(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Combos
    Editor = 0
    ClearChanged_Combo
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ComboEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ItemEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ComboEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Combos
    ClearChanged_Combo
    ClearCombos
    SendRequestCombos
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ComboEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Combo()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Combo_Changed(1), MAX_COMBO * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Combo", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Item()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Item", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Animation.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For I = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(I) = Trim$(.sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For I = 0 To 1
            frmEditor_Animation.scrlSprite(I).Value = .Sprite(I)
            frmEditor_Animation.scrlFrameCount(I).Value = .Frames(I)
            frmEditor_Animation.scrlLoopCount(I).Value = .LoopCount(I)
            
            If .looptime(I) > 0 Then
                frmEditor_Animation.scrlLoopTime(I).Value = .looptime(I)
            Else
                frmEditor_Animation.scrlLoopTime(I).Value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With
    
    Animation_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AnimationEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ANIMATIONS
        If Animation_Changed(I) Then
            Call SendSaveAnimation(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Animation
    Editor = 0
    ClearChanged_Animation
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AnimationEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AnimationEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AnimationEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Animation()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Animation", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim I As Long, mNum As Double
Dim SoundSet As Boolean
Dim Value As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_NPC.Visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .txtName.text = Trim$(NPC(EditorIndex).name)
        .txtAttackSay.text = Trim$(NPC(EditorIndex).AttackSay)
        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.max Then NPC(EditorIndex).Sprite = 0
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        .txtSpawnSecs.text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtChance.text = CStr(NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).DropChance)
        .scrlNum.Value = NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).DropItem
        .scrlValue.Value = NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).DropItemValue
        .txtHP.text = NPC(EditorIndex).HP
        .txtExp.text = NPC(EditorIndex).EXP
        .txtLevel.text = NPC(EditorIndex).Level
        .txtDamage.text = NPC(EditorIndex).Damage
        .scrlAnimation.Value = NPC(EditorIndex).Animation
        If NPC(EditorIndex).speed = 0 Then NPC(EditorIndex).speed = 1
        .scrlMoveSpeed.Value = NPC(EditorIndex).speed
        .chkQuest.Value = NPC(EditorIndex).Quest
        .scrlQuest.Value = NPC(EditorIndex).QuestNum
        .chkRndExp.Value = NPC(EditorIndex).RandExp
        .opPercent_5.Value = CBool(NPC(EditorIndex).Percent_5)
        .opPercent_10.Value = CBool(NPC(EditorIndex).Percent_10)
        .opPercent_20.Value = CBool(NPC(EditorIndex).Percent_20)
        .chkRandHP.Value = NPC(EditorIndex).RandHP
        .txtHPMin.text = NPC(EditorIndex).HPMin
        .chkRandCurrency.Value = NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).RandCurrency
        .opPercent(0).Value = CBool(NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).P_5)
        .opPercent(1).Value = CBool(NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).P_10)
        .opPercent(2).Value = CBool(NPC(EditorIndex).Drops(frmEditor_NPC.scrlDropIndex.Value).P_20)
        .chkRndSpawn.Value = NPC(EditorIndex).RndSpawn
        .txtSpawnSecsMin.text = NPC(EditorIndex).SpawnSecsMin
        If Not .opPercent(0).Value And Not .opPercent(1).Value And Not .opPercent(2).Value Then .opPercent(0).Value = True
        
        If .opPercent_5.Value Then mNum = 0.05
        If .opPercent_10.Value Then mNum = 0.1
        If .opPercent_20.Value Then mNum = 0.2
        
        .lblOutput.Caption = "Variacion Experiencia: " & NPC(EditorIndex).EXP - (NPC(EditorIndex).EXP * mNum) & " - " & NPC(EditorIndex).EXP + (NPC(EditorIndex).EXP * mNum)
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(NPC(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For I = 1 To Stats.Stat_Count - 2
            .scrlStat(I).Value = NPC(EditorIndex).Stat(I)
        Next
        
        For I = 1 To 6
            Select Case I
                Case 1
                    Value = NPC(EditorIndex).Element_Light_Dmg
                Case 2
                    Value = NPC(EditorIndex).Element_Dark_Dmg
                Case 3
                    Value = NPC(EditorIndex).Element_Neut_Dmg
                Case 4
                    Value = NPC(EditorIndex).Element_Light_Res
                Case 5
                    Value = NPC(EditorIndex).Element_Dark_Res
                Case 6
                    Value = NPC(EditorIndex).Element_Neut_Res
            End Select
            
            .scrlElement(I).Value = Value
        Next I
    End With
    
    NPC_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "NpcEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_NPCS
        If NPC_Changed(I) Then
            Call SendSaveNpc(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_NPC
    Editor = 0
    ClearChanged_NPC
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "NpcEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "NpcEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_NPC()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_NPC", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim I As Long
Dim SoundSet As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Resource.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.max = NumResources
        .scrlNormalPic.max = NumResources
        .scrlAnimation.max = MAX_ANIMATIONS
        
        .txtName.text = Trim$(Resource(EditorIndex).name)
        .txtMessage.text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.text = Trim$(Resource(EditorIndex).EmptyMessage)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .txtHPMin.text = Resource(EditorIndex).healthmin
        .txtHPMax.text = Resource(EditorIndex).health
        .chkRandHP.Value = Resource(EditorIndex).HPRand
        .chkDistItems.Value = Resource(EditorIndex).DistItems
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .txtAmountMax.text = Resource(EditorIndex).ItemRewardAmount
        .txtAmountMin.text = Resource(EditorIndex).ItemRewardAmountMin
        .chkRewardRand.Value = Resource(EditorIndex).ItemRewardRand
        .chkSkillExp.Value = Resource(EditorIndex).Exp_Give
        .cmbColor_Success.ListIndex = Resource(EditorIndex).Color_Success
        .cmbColor_Empty.ListIndex = Resource(EditorIndex).Color_Empty
        If Resource(EditorIndex).Exp_Skill > 0 Then
            .cmbSkill.ListIndex = Resource(EditorIndex).Exp_Skill - 1
        Else
            .cmbSkill.ListIndex = 0
        End If
        .cmbSkillReq.ListIndex = Resource(EditorIndex).Skill_Req
        .txtExp.text = Resource(EditorIndex).Exp_Amnt
        .scrlSkillReqLvl.Value = Resource(EditorIndex).Skill_LvlReq
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Resource(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Resource_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ResourceEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_RESOURCES
        If Resource_Changed(I) Then
            Call SendSaveResource(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Resource
    Editor = 0
    ClearChanged_Resource
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ResourceEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResourceEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ResourceEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Resource()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Resource", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Shop.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    
    frmEditor_Shop.txtName.text = Trim$(Shop(EditorIndex).name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "Ninguno"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "Ninguno"

    For I = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem I & ": " & Trim$(Item(I).name)
        frmEditor_Shop.cmbCostItem.AddItem I & ": " & Trim$(Item(I).name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ShopEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim I As Long
Dim Valor As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmEditor_Shop.lstTradeItem.Clear

    For I = 1 To MAX_TRADES
            ' if none, show as none
            If Shop(EditorIndex).TradeItem(I).Item = 0 And Shop(EditorIndex).TradeItem(I).CostItem = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Slot Disponible"
            Else
                 Valor = Shop(EditorIndex).TradeItem(I).CostItem
                If Valor <= "0" Then
                Msgbox "Este objeto no tiene valor alguno"
                Else
                frmEditor_Shop.lstTradeItem.AddItem I & ": " & Shop(EditorIndex).TradeItem(I).ItemValue & "x " & Trim$(Item(Shop(EditorIndex).TradeItem(I).Item).name) & " for " & Shop(EditorIndex).TradeItem(I).CostValue & "x " & Trim$(Item(Shop(EditorIndex).TradeItem(I).CostItem).name)
                End If
            End If

Next
    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UpdateShopTrade", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SHOPS
        If Shop_Changed(I) Then
            Call SendSaveShop(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Shop
    Editor = 0
    ClearChanged_Shop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ShopEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ShopEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Shop()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Shop", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim I As Long
Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Spell.Visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "Ninguno."
    For I = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(I)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.max = MAX_ANIMATIONS
        .scrlAnim.max = MAX_ANIMATIONS
        .scrlAOE.max = MAX_BYTE
        .scrlRange.max = MAX_BYTE
        .scrlMap.max = MAX_MAPS
        .cmdCombatType.ListIndex = 0
        .cmdCombatType.Enabled = False
        .lblCombatType.Caption = "Tipo Combate: "
        .scrlCombatLvl = Spell(EditorIndex).CombatLvlReq
        .scrlCombatLvl.max = MAX_COMBAT_LEVEL
        .scrlCombatLvl.Enabled = False
        .lblCombatLvl.Caption = "Nivel Combate: " & Spell(EditorIndex).CombatLvlReq
        
        ' build class combo
        .cmbClass.Clear
        .cmbClass.AddItem "Ninguno"
        For I = 1 To Max_Classes
            .cmbClass.AddItem Trim$(Class(I).name)
        Next
        
        If Spell(EditorIndex).ClassReq > -1 And Spell(EditorIndex).ClassReq <= Max_Classes Then
            .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        End If
        
        ' set values
        .txtName.text = Trim$(Spell(EditorIndex).name)
        .txtDesc.text = Trim$(Spell(EditorIndex).Desc)
        .cmbType.ListIndex = Spell(EditorIndex).Type
        .scrlMP.Value = Spell(EditorIndex).MPCost
        .scrlLevel.Value = Spell(EditorIndex).LevelReq
        .scrlAccess.Value = Spell(EditorIndex).AccessReq
        .cmbClass.ListIndex = Spell(EditorIndex).ClassReq
        .scrlCast.Value = Spell(EditorIndex).CastTime
        .scrlCool.Value = Spell(EditorIndex).CDTime
        .scrlIcon.Value = Spell(EditorIndex).Icon
        .scrlMap.Value = Spell(EditorIndex).Map
        .scrlX.Value = Spell(EditorIndex).X
        .scrlY.Value = Spell(EditorIndex).Y
        .scrlDir.Value = Spell(EditorIndex).Dir
        .txtVital.text = Spell(EditorIndex).Vital
        .scrlDuration.Value = Spell(EditorIndex).Duration
        .scrlInterval.Value = Spell(EditorIndex).Interval
        .scrlRange.Value = Spell(EditorIndex).Range
        .scrlElement(1).max = MAX_INTEGER
        .scrlElement(2).max = MAX_INTEGER
        .scrlElement(3).max = MAX_INTEGER
        .scrlElement(1).Value = Spell(EditorIndex).Dmg_Light
        .scrlElement(2).Value = Spell(EditorIndex).Dmg_Dark
        .scrlElement(3).Value = Spell(EditorIndex).Dmg_Neut
        .chkparalizar.Value = Spell(EditorIndex).Paralisis 'EaSee Engine 0.5 cielos que bueno soy
        .chkconfusion.Value = Spell(EditorIndex).Inversion
        .chkinvisibilidad.Value = Spell(EditorIndex).Invisibilidad
        .chkveneno.Value = Spell(EditorIndex).Veneno
        .chkVelocidad.Value = Spell(EditorIndex).Velocidad
        .txtcaminar.text = Spell(EditorIndex).VelocidadCaminar2
        .txtcorrer.text = Spell(EditorIndex).VelocidadCorrer2
        .txtfza.text = Spell(EditorIndex).Fuerza
        .txtdes.text = Spell(EditorIndex).Destreza
        .txtagi.text = Spell(EditorIndex).Agilidad
        .txtint.text = Spell(EditorIndex).Inteligencia
        .txtvol.text = Spell(EditorIndex).Voluntad
        .chkbuff.Value = Spell(EditorIndex).Buff
        .chkarco.Value = Spell(EditorIndex).Arco
        .txtarco.text = Spell(EditorIndex).NumeroArcoItem
        .chksprite.Value = Spell(EditorIndex).Sprite
        .txtsprite.text = Spell(EditorIndex).NumeroSprite
        .chktransportar.Value = Spell(EditorIndex).Transportar
        .txtx.text = Spell(EditorIndex).TransportarX
        .txty.text = Spell(EditorIndex).TransportarY
        .txtmapa.text = Spell(EditorIndex).TransportarMapa
        .txtVeneno.text = Spell(EditorIndex).VenenoDmg
        
        If Spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = Spell(EditorIndex).AoE
        .scrlAnimCast.Value = Spell(EditorIndex).CastAnim
        .scrlAnim.Value = Spell(EditorIndex).SpellAnim
        .scrlStun.Value = Spell(EditorIndex).StunDuration
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For I = 0 To .cmbSound.ListCount
                If .cmbSound.List(I) = Trim$(Spell(EditorIndex).sound) Then
                    .cmbSound.ListIndex = I
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Spell_Changed(EditorIndex) = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SpellEditorInit", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorOk(Optional ByVal CloseOut As Boolean = True)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SPELLS
        If Spell_Changed(I) Then
            Call SendSaveSpell(I)
        End If
    Next
    
    If CloseOut Then Unload frmEditor_Spell
    Editor = 0
    ClearChanged_Spell
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SpellEditorOk", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellEditorCancel()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Editor = 0
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SpellEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearChanged_Spell()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearChanged_Spell", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmEditor_Map.fraNpcSpawn.Visible = False
    frmEditor_Map.fraResource.Visible = False
    frmEditor_Map.fraMapItem.Visible = False
    frmEditor_Map.fraMapKey.Visible = False
    frmEditor_Map.fraKeyOpen.Visible = False
    frmEditor_Map.fraMapWarp.Visible = False
    frmEditor_Map.fraShop.Visible = False
    frmEditor_Map.fraSoundEffect.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAttributeDialogue", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


'Event Editor Stuffz Also includes event functions from the map editor (copy/paste/delete)

Sub CopyEvent_Map(X As Long, Y As Long)
Dim count As Long, I As Long
    count = Map.EventCount
    If count = 0 Then Exit Sub
    
    For I = 1 To count
        If Map.Events(I).X = X And Map.Events(I).Y = Y Then
            ' copy it
            'CopyMemory ByVal VarPtr(cpEvent), ByVal VarPtr(Map.Events(i)), LenB(Map.Events(i))
            cpEvent = Map.Events(I)
            ' exit
            Exit Sub
        End If
    Next
End Sub

Sub PasteEvent_Map(X As Long, Y As Long)
Dim count As Long, I As Long, eventNum As Long
    count = Map.EventCount
    
    If count > 0 Then
        For I = 1 To count
            If Map.Events(I).X = X And Map.Events(I).Y = Y Then
                ' already an event - paste over it
                eventNum = I
            End If
        Next
    End If
    
    ' couldn't find one - create one
    If eventNum = 0 Then
        ' increment count
        AddEvent X, Y, True
        eventNum = count + 1
    End If
    
    ' copy it
    'CopyMemory ByVal VarPtr(Map.Events(eventNum)), ByVal VarPtr(cpEvent), LenB(cpEvent)
    Map.Events(eventNum) = cpEvent
    
    ' set position
    Map.Events(eventNum).X = X
    Map.Events(eventNum).Y = Y
End Sub

Sub DeleteEvent(X As Long, Y As Long)
Dim count As Long, I As Long, lowIndex As Long
    If Not InMapEditor Then Exit Sub
    If frmEditor_Events.Visible = True Then Exit Sub
    count = Map.EventCount
    For I = 1 To count
        If Map.Events(I).X = X And Map.Events(I).Y = Y Then
            ' delete it
            ClearEvent I
            lowIndex = I
            Exit For
        End If
    Next
    
    ' not found anything
    If lowIndex = 0 Then Exit Sub
    
    ' move everything down an index
    For I = lowIndex To count - 1
        CopyEvent I + 1, I
    Next
    ' delete the last index
    ClearEvent count
    ' set the new count
    Map.EventCount = count - 1
End Sub

Sub AddEvent(X As Long, Y As Long, Optional ByVal cancelLoad As Boolean = False)
Dim count As Long, pageCount As Long, I As Long
    count = Map.EventCount + 1
    ' make sure there's not already an event
    If count - 1 > 0 Then
        For I = 1 To count - 1
            If Map.Events(I).X = X And Map.Events(I).Y = Y Then
                ' already an event - edit it
                If Not cancelLoad Then EventEditorInit I
                Exit Sub
            End If
        Next
    End If
    ' increment count
    Map.EventCount = count
    ReDim Preserve Map.Events(0 To count)
    ' set the new event
    Map.Events(count).X = X
    Map.Events(count).Y = Y
    ' give it a new page
    pageCount = Map.Events(count).pageCount + 1
    Map.Events(count).pageCount = pageCount
    ReDim Preserve Map.Events(count).Pages(pageCount)
    ' load the editor
    If Not cancelLoad Then EventEditorInit count
End Sub

Sub ClearEvent(eventNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum)))
End Sub

Sub CopyEvent(original As Long, newone As Long)
    CopyMemory ByVal VarPtr(Map.Events(newone)), ByVal VarPtr(Map.Events(original)), LenB(Map.Events(original))
End Sub

Sub EventEditorInit(eventNum As Long)
Dim I As Long
    EditorEvent = eventNum
    ' copy the event data to the temp event
    'CopyMemory ByVal VarPtr(tmpEvent), ByVal VarPtr(Map.Events(eventNum)), LenB(Map.Events(eventNum))
    tmpEvent = Map.Events(eventNum)
    frmEditor_Events.InitEventEditorForm
    ' populate form
    With frmEditor_Events
        ' set the tabs
        .tabPages.Tabs.Clear
        For I = 1 To tmpEvent.pageCount
            .tabPages.Tabs.Add , , str(I)
        Next
        ' items
        .cmbHasItem.Clear
        .cmbHasItem.AddItem "Ninguno"
        For I = 1 To MAX_ITEMS
            .cmbHasItem.AddItem I & ": " & Trim$(Item(I).name)
        Next
        
        ' variables
        .cmbPlayerVar.Clear
        .cmbPlayerVar.AddItem "Ninguno"
        For I = 1 To MAX_VARIABLES
            .cmbPlayerVar.AddItem I & ". " & Variables(I)
        Next
        
        ' variables
        .cmbPlayerSwitch.Clear
        .cmbPlayerSwitch.AddItem "Ninguno"
        For I = 1 To MAX_SWITCHES
            .cmbPlayerSwitch.AddItem I & ". " & Switches(I)
        Next
        
        
        ' name
        .txtName.text = tmpEvent.name
        ' enable delete button
        If tmpEvent.pageCount > 1 Then
            .cmdDeletePage.Enabled = True
        Else
            .cmdDeletePage.Enabled = False
        End If
        .cmdPastePage.Enabled = False
        ' Load page 1 to start off with
        curPageNum = 1
        EventEditorLoadPage curPageNum
    End With
    ' show the editor
    frmEditor_Events.Show
End Sub

Sub EventEditorLoadPage(pageNum As Long)
    ' populate form
    With tmpEvent.Pages(pageNum)
        GraphicSelX = .GraphicX
        GraphicSelY = .GraphicY
        GraphicSelX2 = .GraphicX2
        GraphicSelY2 = .GraphicY2
        frmEditor_Events.cmbGraphic.ListIndex = .GraphicType
        
        frmEditor_Events.cmbHasItem.ListIndex = .HasItemIndex
        frmEditor_Events.txtCondition_itemAmount.text = .HasItemAmount
        frmEditor_Events.cmbMoveFreq.ListIndex = .MoveFreq
        frmEditor_Events.cmbMoveSpeed.ListIndex = .MoveSpeed
        frmEditor_Events.cmbMoveType.ListIndex = .MoveType
        
        frmEditor_Events.cmbPlayerVar.ListIndex = .VariableIndex
        frmEditor_Events.cmbPlayerSwitch.ListIndex = .SwitchIndex
        frmEditor_Events.cmbSelfSwitch.ListIndex = .SelfSwitchIndex
        frmEditor_Events.cmbSelfSwitchCompare.ListIndex = .SelfSwitchCompare
        frmEditor_Events.cmbPlayerVarCompare.ListIndex = .VariableCompare
        
        
        frmEditor_Events.chkGlobal.Value = tmpEvent.Global
        
        frmEditor_Events.cmbTrigger.ListIndex = .Trigger
        frmEditor_Events.chkDirFix.Value = .DirFix
        frmEditor_Events.chkHasItem.Value = .chkHasItem
        frmEditor_Events.chkPlayerVar.Value = .chkVariable
        frmEditor_Events.chkPlayerSwitch.Value = .chkSwitch
        frmEditor_Events.chkSelfSwitch.Value = .chkSelfSwitch
        frmEditor_Events.chkWalkAnim.Value = .WalkAnim
        frmEditor_Events.chkWalkThrough.Value = .Walkthrough
        frmEditor_Events.chkShowName.Value = .ShowName
        frmEditor_Events.txtPlayerVariable = .VariableCondition
        frmEditor_Events.scrlGraphic.Value = .Graphic
        
        If .chkHasItem = 0 Then
            frmEditor_Events.cmbHasItem.Enabled = False
        Else
            frmEditor_Events.cmbHasItem.Enabled = True
        End If
        
        
        If .chkSelfSwitch = 0 Then
            frmEditor_Events.cmbSelfSwitch.Enabled = False
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbSelfSwitch.Enabled = True
            frmEditor_Events.cmbSelfSwitchCompare.Enabled = True
        End If
        
        If .chkSwitch = 0 Then
            frmEditor_Events.cmbPlayerSwitch.Enabled = False
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerSwitch.Enabled = True
            frmEditor_Events.cmbPlayerSwitchCompare.Enabled = True
        End If
        
        
        If .chkVariable = 0 Then
            frmEditor_Events.cmbPlayerVar.Enabled = False
            frmEditor_Events.txtPlayerVariable.Enabled = False
            frmEditor_Events.cmbPlayerVarCompare.Enabled = False
        Else
            frmEditor_Events.cmbPlayerVar.Enabled = True
            frmEditor_Events.txtPlayerVariable.Enabled = True
            frmEditor_Events.cmbPlayerVarCompare.Enabled = True
        End If
        
        If frmEditor_Events.cmbMoveType.ListIndex = 2 Then
            frmEditor_Events.cmdMoveRoute.Enabled = True
        Else
            frmEditor_Events.cmdMoveRoute.Enabled = False
        End If
        
        frmEditor_Events.cmbPositioning.ListIndex = .Position
        
        ' show the commands
        EventListCommands
    End With
End Sub

Sub EventEditorOK()
    ' copy the event data from the temp event
    'CopyMemory ByVal VarPtr(Map.Events(EditorEvent)), ByVal VarPtr(tmpEvent), LenB(tmpEvent)
    Map.Events(EditorEvent) = tmpEvent
    ' unload the form
    Unload frmEditor_Events
End Sub

Public Sub EventListCommands()
Dim I As Long, curlist As Long, oldI As Long, X As Long, indent As String, listleftoff() As Long, conditionalstage() As Long
    frmEditor_Events.lstCommands.Clear
    If tmpEvent.Pages(curPageNum).CommandListCount > 0 Then
    ReDim listleftoff(1 To tmpEvent.Pages(curPageNum).CommandListCount)
    ReDim conditionalstage(1 To tmpEvent.Pages(curPageNum).CommandListCount)
        'Start Up at 1
        curlist = 1
        X = -1
newlist:
        For I = 1 To tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
            If listleftoff(curlist) > 0 Then
                If (tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evCondition Or tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(listleftoff(curlist)).Index = EventType.evShowChoices) And conditionalstage(curlist) <> 0 Then
                    I = listleftoff(curlist)
                ElseIf listleftoff(curlist) >= I Then
                    I = listleftoff(curlist) + 1
                End If
            End If
            If I <= tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
                If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Index = EventType.evCondition Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = I
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Condition
                                Case 0
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Variable [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] != " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data3
                                    End Select
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] == " & "True"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Switch [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1) & "] == " & "False"
                                    End If
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Has Item [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & "] Amount [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2 & "]"
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Class Is [" & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & "]"
                                Case 4
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Knows Skill [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & "]"
                                Case 5
                                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2
                                        Case 0
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                        Case 1
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                        Case 2
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is <= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                        Case 3
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is > " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                        Case 4
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is < " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                        Case 5
                                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player's Level is NOT " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                    End Select
                                Case 6
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2 = 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "True"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "True"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "True"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "True"
                                        End Select
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2 = 1 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [A] == " & "False"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [B] == " & "False"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [C] == " & "False"
                                            Case 3
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Self Switch [D] == " & "False"
                                        End Select
                                    End If
                                Case 7
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 > 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Skill Level [" & Trim$(Skill(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & "] >= " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2
                                    End If
                                Case 8 'Quest status
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1 > 0 Then
                                        Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data2
                                            Case 0
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Quest (" & Trim$(Quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & ") is [Open]"
                                            Case 1
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Quest (" & Trim$(Quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & ") is [Started]"
                                            Case 2
                                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Conditional Branch: Player Quest (" & Trim$(Quest(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.Data1).name) & ") is [Completed]"
                                        End Select
                                    End If
                            
                            End Select
                            
                            indent = indent & "       "
                            listleftoff(curlist) = I
                            conditionalstage(curlist) = 1
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.CommandList
                            GoTo newlist
                        Case 1
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Else"
                            listleftoff(curlist) = I
                            conditionalstage(curlist) = 2
                            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).ConditionalBranch.ElseCommandList
                            GoTo newlist
                        Case 2
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "End Branch"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = I
                            conditionalstage(curlist) = 0
                    End Select
                ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Index = EventType.evShowChoices Then
                    X = X + 1
                    Select Case conditionalstage(curlist)
                        Case 0
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = I
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Show Choices - Prompt: " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "..."
                            
                            indent = indent & "       "
                            listleftoff(curlist) = I
                            conditionalstage(curlist) = 1
                            GoTo newlist
                        Case 1
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text2) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text2) & "]"
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 2
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 2
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 2
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text3) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text3) & "]"
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 3
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 3
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 3
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text4) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text4) & "]"
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 4
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 4
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 4
                            If Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text5) <> "" Then
                                ReDim Preserve EventList(X)
                                EventList(X).CommandList = curlist
                                EventList(X).CommandNum = 0
                                frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "When [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text5) & "]"
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 5
                                curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4
                                GoTo newlist
                            Else
                                X = X - 1
                                listleftoff(curlist) = I
                                conditionalstage(curlist) = 5
                                curlist = curlist
                                GoTo newlist
                            End If
                        Case 5
                            ReDim Preserve EventList(X)
                            EventList(X).CommandList = curlist
                            EventList(X).CommandNum = 0
                            frmEditor_Events.lstCommands.AddItem Mid(indent, 1, Len(indent) - 4) & " : " & "Branch End"
                            indent = Mid(indent, 1, Len(indent) - 7)
                            listleftoff(curlist) = I
                            conditionalstage(curlist) = 0
                    End Select
                Else
                    X = X + 1
                    ReDim Preserve EventList(X)
                    EventList(X).CommandList = curlist
                    EventList(X).CommandNum = I
                    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Index
                        Case EventType.evAddText
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Agregar Texto - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & " - Chat Tipo: Jugador"
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Agregar Texto - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & " - Chat Tipo: Mapa"
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Agregar Texto - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - Color: " & GetColorString(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & " - Chat Tipo: Global"
                            End Select
                        Case EventType.evShowText
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Mostrar Texto - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "..."
                        Case EventType.evPlayerVar
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Variable Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] == " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Variable Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] + " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3
                                Case 2
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Player VariableModificar Variable Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] - " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3
                                Case 3
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Variable Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & Variables(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] Random Between " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & " y " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4
                            End Select
                        Case EventType.evPlayerSwitch
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Switch Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] == True"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Switch Jugador [" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & ". " & Switches(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "] == False"
                            End If
                        Case EventType.evSelfSwitch
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                                Case 0
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [A] a ENCENDIDO"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [A] a APAGADO"
                                    End If
                                Case 1
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [B] a ENCENDIDO"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [B] a APAGADO"
                                    End If
                                Case 2
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [C] a ENCENDIDO"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [C] a APAGADO"
                                    End If
                                Case 3
                                    If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [D] a ENCENDIDO"
                                    ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Switch Propio [D] a APAGADO"
                                    End If
                            End Select
                        Case EventType.evExitProcess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Saliendo del Evento"
                        
                        Case EventType.evChangeItems
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Cantidad de Objeto de [" & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "] a " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Dar " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "(s)"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Quitar " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & " " & Trim$(Item(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "(s) al Jugador."
                            End If
                            
                        Case EventType.evRestoreHP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Regenerar HP"
                        Case EventType.evRestoreMP
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Regenerar MP"
                        Case EventType.evLevelUp
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Subir Nivel"
                        Case EventType.evChangeLevel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Nivel a " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                        Case EventType.evChangeSkills
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Aprender Habilidad [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Olvidar Habilidad [" & Trim$(Spell(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]"
                            End If
                        Case EventType.evChangeClass
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Cambiar a Clase " & Trim$(Class(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name)
                        Case EventType.evChangeSprite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Cambiar a Clase " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                        Case EventType.evChangeSex
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Cambiar a Sexo Masculino."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Cambiar a Sexo Femenino."
                            End If
                        Case EventType.evChangePK
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "PK modificado a No."
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "PK modificado a Si."
                            End If
                        Case EventType.evWarpPlayer
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Transportar a Mapa: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & ") conservando direccion."
                            Else
                                Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4 - 1
                                    Case DIR_UP
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Transportar al Mapa: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & ") mirando arriba."
                                    Case DIR_DOWN
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Transportar al Mapa: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & ") mirando abajo."
                                    Case DIR_LEFT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Transportar al Mapa: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & ") mirando izquierda."
                                    Case DIR_RIGHT
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Transportar al Mapa: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & ") mirando derecha."
                                End Select
                            End If
                        Case EventType.evSetMoveRoute
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 <= Map.EventCount Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar la ruta del evento #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]"
                            Else
                               frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar ruta de EVENTO NO ENCONTRADO!"
                            End If
                        Case EventType.evPlayAnimation
                            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Reprod Animacion " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]" & " en el Jugador"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 1 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Reprod Animacion " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]" & " en el Evento #" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & " [" & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3).name) & "]"
                            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2 = 2 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Reprod Animacion " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " [" & Trim$(Animation(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]" & " en el Tile(" & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3 & "," & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4 & ")"
                            End If
                        Case EventType.evCustomScript
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Ejecutar Script Numero: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                        Case EventType.evPlayBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "BGM [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1) & "]"
                        Case EventType.evFadeoutBGM
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Apagar BGM"
                        Case EventType.evPlaySound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "SFX [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1) & "]"
                        Case EventType.evStopSound
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Apagar SFX"
                        Case EventType.evOpenBank
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Abrir Banco"
                        Case EventType.evOpenShop
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Abrir Tienda [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & ". " & Trim$(Shop(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1).name) & "]"
                        Case EventType.evSetAccess
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Modificar Privilegios [" & frmEditor_Events.cmbSetAccess.List(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "]"
                        Case EventType.evGiveExp
                            Select Case Abs(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2)
                                Case 1
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Dar " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " Experiencia de Combate."
                                Case 0
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Dar " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1 & " [" & Trim$(Skill(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3).name) & "] Exp."
                            End Select
                        Case EventType.evShowChatBubble
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                                Case TARGET_TYPE_PLAYER
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Burbuja de Chat - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - En Jugador"
                                Case TARGET_TYPE_NPC
                                    If Map.NPC(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) <= 0 Then
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Burbuja de Chat - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - En NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & ". ]"
                                    Else
                                        frmEditor_Events.lstCommands.AddItem indent & "@>" & "Burbuja de Chat - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - En NPC [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & ". " & Trim$(NPC(Map.NPC(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2)).name) & "]"
                                    End If
                                Case TARGET_TYPE_EVENT
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Burbuja de Chat - " & Mid(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1, 1, 20) & "... - En Evento [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & ". " & Trim$(Map.Events(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2).name) & "]"
                            End Select
                        Case EventType.evLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Etiqueta: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1) & "]"
                        Case EventType.evGotoLabel
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Ir a Etiqueta: [" & Trim$(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Text1) & "]"
                        Case EventType.evSpawnNpc
                            If Map.NPC(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) <= 0 Then
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawnear NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & ". " & "]"
                            Else
                                frmEditor_Events.lstCommands.AddItem indent & "@>" & "Spawnear NPC: [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & ". " & Trim$(NPC(Map.NPC(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1)).name) & "]"
                            End If
                        Case EventType.evFadeIn
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade In"
                        Case EventType.evFadeOut
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Fade Out"
                        Case EventType.evFlashWhite
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Flash Blanco"
                        Case EventType.evSetFog
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Niebla [Niebla: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & " Velocidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & " Opacidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3) & "]"
                        Case EventType.evSetWeather
                            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1
                                Case WEATHER_TYPE_NONE
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Clima [Ninguno]"
                                Case WEATHER_TYPE_RAIN
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Clima [Lluvia - Intensidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_HAIL
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Clima [Granizo - Intensidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_SNOW
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Nieve - Intensidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_SANDSTORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Torm Arena - Intensidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "]"
                                Case WEATHER_TYPE_STORM
                                    frmEditor_Events.lstCommands.AddItem indent & "@>" & "Set Weather [Tormenta - Intensidad: " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "]"
                            End Select
                        Case EventType.evSetTint
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Matiz Mapa RGBA [" & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data2) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data3) & "," & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data4) & "]"
                        Case EventType.evWait
                            frmEditor_Events.lstCommands.AddItem indent & "@>" & "Esperar " & CStr(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I).Data1) & " Ms"
                        Case Else
                            'Ghost
                            X = X - 1
                            If X = -1 Then
                                ReDim EventList(0)
                            Else
                                ReDim Preserve EventList(X)
                            End If
                    End Select
                End If
            End If
        Next
        If curlist > 1 Then
            X = X + 1
            ReDim Preserve EventList(X)
            EventList(X).CommandList = curlist
            EventList(X).CommandNum = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
            frmEditor_Events.lstCommands.AddItem indent & "@> "
            curlist = tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList
            GoTo newlist
        End If
    End If
    
    frmEditor_Events.lstCommands.AddItem indent & "@> "
End Sub

Sub ListCommandAdd(s As String)
Static X As Long
    frmEditor_Events.lstCommands.AddItem s
    ' scrollbar
    If X < frmEditor_Events.TextWidth(s & "  ") Then
       X = frmEditor_Events.TextWidth(s & "  ")
      If frmEditor_Events.ScaleMode = vbTwips Then X = X / Screen.TwipsPerPixelX ' if twips change to pixels
      SendMessageByNum frmEditor_Events.lstCommands.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    End If
End Sub

Sub AddCommand(Index As Long)
    Dim curlist As Long, I As Long, X As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curlist = 1
    Else
        curlist = EventList(frmEditor_Events.lstCommands.ListIndex).CommandList
    End If
        
    If tmpEvent.Pages(curPageNum).CommandListCount = 0 Then
        tmpEvent.Pages(curPageNum).CommandListCount = 1
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist)
    End If
    
    oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
    tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount + 1
    p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    If p <= 0 Then
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
    Else
        ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(1 To p)
        tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
        For I = 1 To p - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(I) = oldCommandList.Commands(I)
        Next
    End If
    
    If frmEditor_Events.lstCommands.ListIndex = frmEditor_Events.lstCommands.ListCount - 1 Then
        curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
    Else
        I = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        If I < tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
            For X = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1 To I Step -1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X + 1) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X)
            Next
            curslot = EventList(frmEditor_Events.lstCommands.ListIndex).CommandNum
        Else
            curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        End If
    End If
    
    
    Select Case Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            'This is the part where the whole entire source goes to hell :D
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 2
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.CommandList).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.ElseCommandList).ParentList = curlist
            For I = 0 To 8
                If frmEditor_Events.optCondition_Index(I).Value = True Then X = I
            Next
            
            Select Case X
                Case 0 'Player Var
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
                Case 1 'Player Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
                Case 2 'Has Item
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = Val(frmEditor_Events.txtCondition_itemAmount.text)
                Case 3 'Class Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
                Case 4 'Learnt Skill
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
                Case 5 'Level Is
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
                Case 6 'Self Switch
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
                Case 7 'Skill Level
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 7
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SkillReq.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = Val(frmEditor_Events.txtCondition_SkillLvlReq.text)
                Case 8 'Quest Status
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 8
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_Quest.ListIndex + 1
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_Status.ListIndex
            End Select
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
            
            tmpEvent.Pages(curPageNum).CommandListCount = tmpEvent.Pages(curPageNum).CommandListCount + 4
            ReDim Preserve tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = tmpEvent.Pages(curPageNum).CommandListCount - 3
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = tmpEvent.Pages(curPageNum).CommandListCount - 2
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = tmpEvent.Pages(curPageNum).CommandListCount - 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = tmpEvent.Pages(curPageNum).CommandListCount
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 3).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 2).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount - 1).ParentList = curlist
            tmpEvent.Pages(curPageNum).CommandList(tmpEvent.Pages(curPageNum).CommandListCount).ParentList = curlist
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For I = 0 To 3
                If frmEditor_Events.optVariableAction(I).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = I
            If I = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(I).text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(I + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(I).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evExitProcess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evRestoreHP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evRestoreMP
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evLevelUp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evFadeoutBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evStopSound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenBank
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.opMine.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbSkilling.ListIndex + 1
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evFadeIn
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFadeOut
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evFlashWhite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index = Index
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.Value
    End Select
    EventListCommands
End Sub

Public Sub EditEventCommand()
    Dim I As Long, X As Long, Z As Long, curlist As Long, curslot As Long
    I = frmEditor_Events.lstCommands.ListIndex
    If I = -1 Then Exit Sub
    
    If I > UBound(EventList) Then Exit Sub

    curlist = EventList(I).CommandList
    curslot = EventList(I).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            isEdit = True
            frmEditor_Events.txtAddText_Text.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.scrlAddText_Colour.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
                Case 0
                    frmEditor_Events.optAddText_Player.Value = True
                Case 1
                    frmEditor_Events.optAddText_Map.Value = True
                Case 2
                    frmEditor_Events.optAddText_Global.Value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(2).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCondition
            isEdit = True
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(7).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.ClearConditionFrame
            frmEditor_Events.optCondition_Index(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition).Value = True
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition
                Case 0
                    frmEditor_Events.cmbCondition_PlayerVarIndex.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarCompare.Enabled = True
                    frmEditor_Events.txtCondition_PlayerVarCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                    frmEditor_Events.txtCondition_PlayerVarCondition.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3
                Case 1
                    frmEditor_Events.cmbCondition_PlayerSwitch.Enabled = True
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 2
                    frmEditor_Events.cmbCondition_HasItem.Enabled = True
                    frmEditor_Events.txtCondition_itemAmount.Enabled = True
                    frmEditor_Events.cmbCondition_HasItem.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.txtCondition_itemAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 3
                    frmEditor_Events.cmbCondition_ClassIs.Enabled = True
                    frmEditor_Events.cmbCondition_ClassIs.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 4
                    frmEditor_Events.cmbCondition_LearntSkill.Enabled = True
                    frmEditor_Events.cmbCondition_LearntSkill.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                Case 5
                    frmEditor_Events.cmbCondition_LevelCompare.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.Enabled = True
                    frmEditor_Events.txtCondition_LevelAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_LevelCompare.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 6
                    frmEditor_Events.cmbCondition_SelfSwitch.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.Enabled = True
                    frmEditor_Events.cmbCondition_SelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1
                    frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 7
                    frmEditor_Events.cmbCondition_SkillReq.Enabled = True
                    frmEditor_Events.txtCondition_SkillLvlReq.Enabled = True
                    frmEditor_Events.cmbCondition_SkillReq.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.txtCondition_SkillLvlReq.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
                Case 8 'Quest Status
                    frmEditor_Events.cmbCondition_Quest.Enabled = True
                    frmEditor_Events.cmbCondition_Status.Enabled = True
                    frmEditor_Events.cmbCondition_Quest.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 - 1
                    frmEditor_Events.cmbCondition_Status.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2
            End Select
        Case EventType.evShowText
            isEdit = True
            frmEditor_Events.txtShowText.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(0).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChoices
            isEdit = True
            frmEditor_Events.txtChoicePrompt.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.txtChoices(1).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2
            frmEditor_Events.txtChoices(2).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3
            frmEditor_Events.txtChoices(3).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4
            frmEditor_Events.txtChoices(4).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(1).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerVar
            isEdit = True
            frmEditor_Events.cmbVariable.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
                Case 0
                    frmEditor_Events.optVariableAction(0).Value = True
                    frmEditor_Events.txtVariableData(0).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 1
                    frmEditor_Events.optVariableAction(1).Value = True
                    frmEditor_Events.txtVariableData(1).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 2
                    frmEditor_Events.optVariableAction(2).Value = True
                    frmEditor_Events.txtVariableData(2).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                Case 3
                    frmEditor_Events.optVariableAction(3).Value = True
                    frmEditor_Events.txtVariableData(3).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                    frmEditor_Events.txtVariableData(4).text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(4).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayerSwitch
            isEdit = True
            frmEditor_Events.cmbSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayerSwitchSet.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(5).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSelfSwitch
            isEdit = True
            frmEditor_Events.cmbSetSelfSwitch.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.cmbSetSelfSwitchTo.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(6).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeItems
            isEdit = True
            frmEditor_Events.cmbChangeItemIndex.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeItemSet.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeItemAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optChangeItemRemove.Value = True
            End If
            frmEditor_Events.txtChangeItemsAmount.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(10).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeLevel
            isEdit = True
            frmEditor_Events.scrlChangeLevel.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(11).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSkills
            isEdit = True
            frmEditor_Events.cmbChangeSkills.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optChangeSkillsAdd.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optChangeSkillsRemove.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(12).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeClass
            isEdit = True
            frmEditor_Events.cmbChangeClass.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(13).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSprite
            isEdit = True
            frmEditor_Events.scrlChangeSprite.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(14).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangeSex
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangeSexMale.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangeSexFemale.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(15).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evChangePK
            isEdit = True
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1 Then
                frmEditor_Events.optChangePKYes.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0 Then
                frmEditor_Events.optChangePKNo.Value = True
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(16).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWarpPlayer
            isEdit = True
            frmEditor_Events.scrlWPMap.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWPX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.scrlWPY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.cmbWarpPlayerDir.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(18).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetMoveRoute
            isEdit = True
            frmEditor_Events.fraMoveRoute.Visible = True
            frmEditor_Events.lstMoveRoute.Clear
            frmEditor_Events.cmbEvent.Clear
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            frmEditor_Events.cmbEvent.AddItem "This Event"
            frmEditor_Events.cmbEvent.ListIndex = 0
            frmEditor_Events.cmbEvent.Enabled = True
            
            For I = 1 To Map.EventCount
                If I <> EditorEvent Then
                    frmEditor_Events.cmbEvent.AddItem Trim$(Map.Events(I).name)
                    X = X + 1
                    ListOfEvents(X) = I
                    If I = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 Then frmEditor_Events.cmbEvent.ListIndex = X
                End If
            Next
            
                
            IsMoveRouteCommand = True
                
            frmEditor_Events.chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                
            TempMoveRouteCount = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount
            TempMoveRoute = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute
            
            For I = 1 To TempMoveRouteCount
                Select Case TempMoveRoute(I).Index
                    Case 1
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Arriba"
                    Case 2
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Abajo"
                    Case 3
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Izquierda"
                    Case 4
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Derecha"
                    Case 5
                        frmEditor_Events.lstMoveRoute.AddItem "Mover al Azar"
                    Case 6
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Hacia Jugador"
                    Case 7
                        frmEditor_Events.lstMoveRoute.AddItem "Mover Lejos Jugador"
                    Case 8
                        frmEditor_Events.lstMoveRoute.AddItem "Paso Adelante"
                    Case 9
                        frmEditor_Events.lstMoveRoute.AddItem "Paso Atras"
                    Case 10
                        frmEditor_Events.lstMoveRoute.AddItem "Esperar 100ms"
                    Case 11
                        frmEditor_Events.lstMoveRoute.AddItem "Esperar 500ms"
                    Case 12
                        frmEditor_Events.lstMoveRoute.AddItem "Esperar 1000ms"
                    Case 13
                        frmEditor_Events.lstMoveRoute.AddItem "Mirar Arriba"
                    Case 14
                        frmEditor_Events.lstMoveRoute.AddItem "Mirar Abajo"
                    Case 15
                        frmEditor_Events.lstMoveRoute.AddItem "Mirar Izq"
                    Case 16
                        frmEditor_Events.lstMoveRoute.AddItem "Mirar Der"
                    Case 17
                        frmEditor_Events.lstMoveRoute.AddItem "Girar 90° a la Derecha"
                    Case 18
                        frmEditor_Events.lstMoveRoute.AddItem "Girar 90° a la Izquierda"
                    Case 19
                        frmEditor_Events.lstMoveRoute.AddItem "Girar 180°"
                    Case 20
                        frmEditor_Events.lstMoveRoute.AddItem "Girar al Azar"
                    Case 21
                        frmEditor_Events.lstMoveRoute.AddItem "Girar Hacia Jugador"
                    Case 22
                        frmEditor_Events.lstMoveRoute.AddItem "Girar Contra Jugador"
                    Case 23
                        frmEditor_Events.lstMoveRoute.AddItem "Ralentizar 8x"
                    Case 24
                        frmEditor_Events.lstMoveRoute.AddItem "Ralentizar 4x"
                    Case 25
                        frmEditor_Events.lstMoveRoute.AddItem "Ralentizar 2x"
                    Case 26
                        frmEditor_Events.lstMoveRoute.AddItem "Velocidad Normal"
                    Case 27
                        frmEditor_Events.lstMoveRoute.AddItem "Acelerar 2x"
                    Case 28
                        frmEditor_Events.lstMoveRoute.AddItem "Acelerar 4x"
                    Case 29
                        frmEditor_Events.lstMoveRoute.AddItem "Frecuencia Minima"
                    Case 30
                        frmEditor_Events.lstMoveRoute.AddItem "Frecuencia Menor"
                    Case 31
                        frmEditor_Events.lstMoveRoute.AddItem "Frecuencia Normal"
                    Case 32
                        frmEditor_Events.lstMoveRoute.AddItem "Frecuencia Mayor"
                    Case 33
                        frmEditor_Events.lstMoveRoute.AddItem "Frecuencia Maxima"
                    Case 34
                        frmEditor_Events.lstMoveRoute.AddItem "Animacion al Caminar"
                    Case 35
                        frmEditor_Events.lstMoveRoute.AddItem "Sin Animacion"
                    Case 36
                        frmEditor_Events.lstMoveRoute.AddItem "Corregir Direccion"
                    Case 37
                        frmEditor_Events.lstMoveRoute.AddItem "No Corregir Direccion"
                    Case 38
                        frmEditor_Events.lstMoveRoute.AddItem "Pasar a Traves"
                    Case 39
                        frmEditor_Events.lstMoveRoute.AddItem "No Atravezar"
                    Case 40
                        frmEditor_Events.lstMoveRoute.AddItem "Posicion Debajo de Jugador"
                    Case 41
                        frmEditor_Events.lstMoveRoute.AddItem "Posicion Capa de Jugador"
                    Case 42
                        frmEditor_Events.lstMoveRoute.AddItem "Posicion Sobre Jugador"
                    Case 43
                        frmEditor_Events.lstMoveRoute.AddItem "Sprite"
                End Select
            Next
                
            frmEditor_Events.fraMoveRoute.Width = 841
            frmEditor_Events.fraMoveRoute.Height = 609
            frmEditor_Events.fraMoveRoute.Visible = True
            
            frmEditor_Events.fraDialogue.Visible = False
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayAnimation
            isEdit = True
            frmEditor_Events.lblPlayAnimX.Visible = False
            frmEditor_Events.lblPlayAnimY.Visible = False
            frmEditor_Events.scrlPlayAnimTileX.Visible = False
            frmEditor_Events.scrlPlayAnimTileY.Visible = False
            frmEditor_Events.cmbPlayAnimEvent.Visible = False
            frmEditor_Events.cmbPlayAnim.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.cmbPlayAnimEvent.Clear
            For I = 1 To Map.EventCount
                frmEditor_Events.cmbPlayAnimEvent.AddItem I & ". " & Trim$(Map.Events(I).name)
            Next
            frmEditor_Events.cmbPlayAnimEvent.ListIndex = 0
            If tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0 Then
                frmEditor_Events.optPlayAnimPlayer.Value = True
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1 Then
                frmEditor_Events.optPlayAnimEvent.Value = True
                frmEditor_Events.cmbPlayAnimEvent.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 - 1
            ElseIf tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2 Then
                frmEditor_Events.optPlayAnimTile.Value = True
                frmEditor_Events.scrlPlayAnimTileX.max = Map.MaxX
                frmEditor_Events.scrlPlayAnimTileY.max = Map.MaxY
                frmEditor_Events.scrlPlayAnimTileX.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
                frmEditor_Events.scrlPlayAnimTileY.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            End If
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(20).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evCustomScript
            isEdit = True
            frmEditor_Events.scrlCustomScript.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(29).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlayBGM
            isEdit = True
            For I = 1 To UBound(musicCache())
                If musicCache(I) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlayBGM.ListIndex = I - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(25).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evPlaySound
            isEdit = True
            For I = 1 To UBound(soundCache())
                If soundCache(I) = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 Then
                    frmEditor_Events.cmbPlaySound.ListIndex = I - 1
                End If
            Next
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(26).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evOpenShop
            isEdit = True
            frmEditor_Events.cmbOpenShop.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(21).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetAccess
            isEdit = True
            frmEditor_Events.cmbSetAccess.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(28).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGiveExp
            isEdit = True
            frmEditor_Events.scrlGiveExp.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.lblGiveExp.Caption = "Dar Exp: " & tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.opMine.Value = CBool(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2)
            frmEditor_Events.opSkilling.Value = Not CBool(tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2)
            frmEditor_Events.cmbSkilling.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(17).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evShowChatBubble
            isEdit = True
            frmEditor_Events.txtChatbubbleText.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
                Case TARGET_TYPE_PLAYER
                    frmEditor_Events.optChatBubbleTarget(0).Value = True
                Case TARGET_TYPE_NPC
                    frmEditor_Events.optChatBubbleTarget(1).Value = True
                Case TARGET_TYPE_EVENT
                    frmEditor_Events.optChatBubbleTarget(2).Value = True
            End Select
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(3).Visible = True
            frmEditor_Events.fraCommands.Visible = False
            frmEditor_Events.cmbChatBubbleTarget.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 - 1
        Case EventType.evLabel
            isEdit = True
            frmEditor_Events.txtLabelName.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(8).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evGotoLabel
            isEdit = True
            frmEditor_Events.txtGotoLabel.text = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(9).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSpawnNpc
            isEdit = True
            frmEditor_Events.cmbSpawnNPC.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 - 1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(19).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetFog
            isEdit = True
            frmEditor_Events.ScrlFogData(0).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.ScrlFogData(1).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.ScrlFogData(2).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(22).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetWeather
            isEdit = True
            frmEditor_Events.CmbWeather.ListIndex = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlWeatherIntensity.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(23).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evSetTint
            isEdit = True
            frmEditor_Events.scrlMapTintData(0).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.scrlMapTintData(1).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2
            frmEditor_Events.scrlMapTintData(2).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3
            frmEditor_Events.scrlMapTintData(3).Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(24).Visible = True
            frmEditor_Events.fraCommands.Visible = False
        Case EventType.evWait
            isEdit = True
            frmEditor_Events.scrlWaitAmount.Value = tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1
            frmEditor_Events.fraDialogue.Visible = True
            frmEditor_Events.fraCommand(27).Visible = True
            frmEditor_Events.fraCommands.Visible = False
    End Select
End Sub

Public Sub DeleteEventCommand()
    Dim I As Long, X As Long, Z As Long, curlist As Long, curslot As Long, p As Long, oldCommandList As CommandListRec
    I = frmEditor_Events.lstCommands.ListIndex
    If I = -1 Then Exit Sub
    
    If I > UBound(EventList) Then Exit Sub
    
    curlist = EventList(I).CommandList
    curslot = EventList(I).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    If curslot = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            X = 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For I = 1 To p + 1
                If I <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(I)
                    X = X + 1
                End If
            Next
        End If
    Else
        tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount - 1
        p = tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount
        oldCommandList = tmpEvent.Pages(curPageNum).CommandList(curlist)
        X = 1
        If p <= 0 Then
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(0)
        Else
            ReDim tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(p)
            tmpEvent.Pages(curPageNum).CommandList(curlist).ParentList = oldCommandList.ParentList
            tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount = p
            For I = 1 To p + 1
                If I <> curslot Then
                    tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(X) = oldCommandList.Commands(I)
                    X = X + 1
                End If
            Next
        End If
    End If
    
    EventListCommands
    
End Sub

Public Sub ClearEventCommands()
    ReDim tmpEvent.Pages(curPageNum).CommandList(1)
    tmpEvent.Pages(curPageNum).CommandListCount = 1
    EventListCommands
End Sub

Public Sub EditCommand()
    Dim I As Long, X As Long, Z As Long, curlist As Long, curslot As Long
    I = frmEditor_Events.lstCommands.ListIndex
    If I = -1 Then Exit Sub
    
    If I > UBound(EventList) Then Exit Sub

    curlist = EventList(I).CommandList
    curslot = EventList(I).CommandNum
    
    If curlist = 0 Then Exit Sub
    If curslot = 0 Then Exit Sub
    
    If curlist > tmpEvent.Pages(curPageNum).CommandListCount Then Exit Sub
    If curslot > tmpEvent.Pages(curPageNum).CommandList(curlist).CommandCount Then Exit Sub
    
    Select Case tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Index
        Case EventType.evAddText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtAddText_Text.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlAddText_Colour.Value
            If frmEditor_Events.optAddText_Player.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optAddText_Map.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optAddText_Global.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
        Case EventType.evCondition
            If frmEditor_Events.optCondition_Index(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 0
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerVarIndex.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_PlayerVarCompare.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data3 = Val(frmEditor_Events.txtCondition_PlayerVarCondition.text)
            ElseIf frmEditor_Events.optCondition_Index(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_PlayerSwitch.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondtion_PlayerSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_HasItem.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = Val(frmEditor_Events.txtCondition_itemAmount.text)
            ElseIf frmEditor_Events.optCondition_Index(3).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 3
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_ClassIs.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(4).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 4
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_LearntSkill.ListIndex + 1
            ElseIf frmEditor_Events.optCondition_Index(5).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 5
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = Val(frmEditor_Events.txtCondition_LevelAmount.text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_LevelCompare.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(6).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 6
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SelfSwitch.ListIndex
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_SelfSwitchCondition.ListIndex
            ElseIf frmEditor_Events.optCondition_Index(7).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 7
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_SkillReq.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.txtCondition_SkillLvlReq.text
            ElseIf frmEditor_Events.optCondition_Index(8).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Condition = 8
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data1 = frmEditor_Events.cmbCondition_Quest.ListIndex + 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).ConditionalBranch.Data2 = frmEditor_Events.cmbCondition_Status.ListIndex
            End If
        Case EventType.evShowText
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtShowText.text
        Case EventType.evShowChoices
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChoicePrompt.text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text2 = frmEditor_Events.txtChoices(1).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text3 = frmEditor_Events.txtChoices(2).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text4 = frmEditor_Events.txtChoices(3).text
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text5 = frmEditor_Events.txtChoices(4).text
        Case EventType.evPlayerVar
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbVariable.ListIndex + 1
            For I = 0 To 3
                If frmEditor_Events.optVariableAction(I).Value = True Then
                    Exit For
                End If
            Next
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = I
            If I = 3 Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(I).text)
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = Val(frmEditor_Events.txtVariableData(I + 1).text)
            Else
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtVariableData(I).text)
            End If
        Case EventType.evPlayerSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSwitch.ListIndex + 1
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbPlayerSwitchSet.ListIndex
        Case EventType.evSelfSwitch
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetSelfSwitch.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbSetSelfSwitchTo.ListIndex
        Case EventType.evChangeItems
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeItemIndex.ListIndex + 1
            If frmEditor_Events.optChangeItemSet.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeItemAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            ElseIf frmEditor_Events.optChangeItemRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
            End If
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = Val(frmEditor_Events.txtChangeItemsAmount.text)
        Case EventType.evChangeLevel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeLevel.Value
        Case EventType.evChangeSkills
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeSkills.ListIndex + 1
            If frmEditor_Events.optChangeSkillsAdd.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optChangeSkillsRemove.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
            End If
        Case EventType.evChangeClass
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbChangeClass.ListIndex + 1
        Case EventType.evChangeSprite
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlChangeSprite.Value
        Case EventType.evChangeSex
            If frmEditor_Events.optChangeSexMale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            ElseIf frmEditor_Events.optChangeSexFemale.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            End If
        Case EventType.evChangePK
            If frmEditor_Events.optChangePKYes.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 1
            ElseIf frmEditor_Events.optChangePKNo.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = 0
            End If
        Case EventType.evWarpPlayer
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWPMap.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWPX.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlWPY.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.cmbWarpPlayerDir.ListIndex
        Case EventType.evSetMoveRoute
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = ListOfEvents(frmEditor_Events.cmbEvent.ListIndex)
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.chkIgnoreMove.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.chkRepeatRoute.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRouteCount = TempMoveRouteCount
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).MoveRoute = TempMoveRoute
        Case EventType.evPlayAnimation
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbPlayAnim.ListIndex + 1
            If frmEditor_Events.optPlayAnimPlayer.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 0
            ElseIf frmEditor_Events.optPlayAnimEvent.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 1
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbPlayAnimEvent.ListIndex + 1
            ElseIf frmEditor_Events.optPlayAnimTile.Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = 2
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlPlayAnimTileX.Value
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlPlayAnimTileY.Value
            End If
        Case EventType.evCustomScript
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlCustomScript.Value
        Case EventType.evPlayBGM
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = musicCache(frmEditor_Events.cmbPlayBGM.ListIndex + 1)
        Case EventType.evPlaySound
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = soundCache(frmEditor_Events.cmbPlaySound.ListIndex + 1)
        Case EventType.evOpenShop
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbOpenShop.ListIndex + 1
        Case EventType.evSetAccess
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSetAccess.ListIndex
        Case EventType.evGiveExp
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlGiveExp.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.opMine.Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.cmbSkilling.ListIndex + 1
        Case EventType.evShowChatBubble
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtChatbubbleText.text
            If frmEditor_Events.optChatBubbleTarget(0).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_PLAYER
            ElseIf frmEditor_Events.optChatBubbleTarget(1).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_NPC
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            ElseIf frmEditor_Events.optChatBubbleTarget(2).Value = True Then
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = TARGET_TYPE_EVENT
                tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.cmbChatBubbleTarget.ListIndex + 1
            End If
        Case EventType.evLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtLabelName.text
        Case EventType.evGotoLabel
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Text1 = frmEditor_Events.txtGotoLabel.text
        Case EventType.evSpawnNpc
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.cmbSpawnNPC.ListIndex + 1
        Case EventType.evSetFog
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.ScrlFogData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.ScrlFogData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.ScrlFogData(2).Value
        Case EventType.evSetWeather
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.CmbWeather.ListIndex
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlWeatherIntensity.Value
        Case EventType.evSetTint
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlMapTintData(0).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data2 = frmEditor_Events.scrlMapTintData(1).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data3 = frmEditor_Events.scrlMapTintData(2).Value
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data4 = frmEditor_Events.scrlMapTintData(3).Value
        Case EventType.evWait
            tmpEvent.Pages(curPageNum).CommandList(curlist).Commands(curslot).Data1 = frmEditor_Events.scrlWaitAmount.Value
    End Select
    
    EventListCommands
End Sub

Sub RequestSwitchesAndVariables()
Dim I As Long, buffer As clsBuffer
Set buffer = New clsBuffer
buffer.WriteLong CRequestSwitchesAndVariables
SendData buffer.ToArray
Set buffer = Nothing
End Sub

Sub SendSwitchesAndVariables()
Dim I As Long, buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CSwitchesAndVariables
    For I = 1 To MAX_SWITCHES
        buffer.WriteString Switches(I)
    Next
    For I = 1 To MAX_VARIABLES
        buffer.WriteString Variables(I)
    Next
    SendData buffer.ToArray
Set buffer = Nothing
End Sub




Public Sub ActualizarMapaCubos()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong CNeedMap
    buffer.WriteLong 1
    SendData buffer.ToArray()
    RedibujarMapaCubos
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "MapEditorCancel", "modGameEditors", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub






