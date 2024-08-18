Attribute VB_Name = "modServerTCP"
Option Explicit
Public MDE As String

Sub UpdateCaption()
    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"
End Sub

Sub CreateFullMapCache()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call MapCache_Create(I)
    Next

End Sub

Function IsConnected(ByVal index As Long) As Boolean

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

End Function

Function isPlaying(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            isPlaying = True
        End If
    End If

End Function

Function IsLoggedIn(ByVal index As Long) As Boolean

    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If LCase$(Trim$(Player(I).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If

    Next

End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim I As Long
    Dim N As Long

    For I = 1 To Player_HighIndex

        If IsConnected(I) Then
            If Trim$(GetPlayerIP(I)) = IP Then
                N = N + 1

                If (N > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If

    Next

End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long
    filename = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData buffer.ToArray()
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            Call SendDataTo(I, Data)
        End If

    Next

End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If I <> index Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If GetPlayerMap(I) = mapnum Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub ClearFList(ByVal index As Long)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SClearFList
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFriendsList(ByVal index As Long)
Dim buffer As clsBuffer
Dim I As Long, tempName As String
    
    For I = 1 To GetPlayerFriends(index)
        Set buffer = New clsBuffer
        
        If isPlaying(FindPlayer(GetPlayerFriendName(index, I))) Then
            tempName = GetPlayerFriendName(index, I) & " (Online)"
        Else
            tempName = GetPlayerFriendName(index, I) & " (Offline)"
        End If
        
        buffer.WriteLong SUpdateFList
        buffer.WriteString tempName
        SendDataTo index, buffer.ToArray
        Set buffer = Nothing
    Next I
End Sub

Sub SendUpdateFriendsLists()
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            Call ClearFList(I)
            Call UpdateFriendsList(I)
        End If

    Next
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapnum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If GetPlayerMap(I) = mapnum Then
                If I <> index Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If

    Next

End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
Dim I As Long

    For I = 1 To Party(partyNum).MemberCount
        If Party(partyNum).Member(I) > 0 Then
            Call SendDataTo(Party(partyNum).Member(I), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Dim I As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteLong Color

    For I = 1 To Player_HighIndex
        If isPlaying(I) And GetPlayerAccess(I) > 0 Then
            SendDataTo I, buffer.ToArray
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub SendPlayerFollow(ByVal index As Long, ByVal Dir As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SFollowPlayer
    buffer.WriteByte Dir
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal Color As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToMap mapnum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    
    Set buffer = Nothing
End Sub
Public Sub SendAchievement(ByVal index As Long, ByVal Msg As String, personal As Boolean)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    If personal = True Then
    Msg = "personal-" & Msg
    Else
    Msg = "total-" & Msg
    End If
    
    buffer.WriteLong SSendAchievement
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray
    DoEvents
    
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal Color As Byte)
Dim I As Long
    ' send message to all people
    For I = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(I) > 0 Then
            ' make sure they're logged on
            If IsConnected(Party(partyNum).Member(I)) And isPlaying(Party(partyNum).Member(I)) Then
                PlayerMsg Party(partyNum).Member(I), Msg, Color
            End If
        End If
    Next
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
Dim var1 As String

    If index > 0 Then
        If isPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " kicked (" & Reason & ")", White)
        End If

                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M5")
        Call AlertMsg(index, var1 & Options.Game_Name & ".")
    End If      '[SERVIDOR]M5

End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim I As Long

    If (index = 0) Then
        I = FindOpenPlayerSlot

        If I <> 0 Then
            ' we can connect them
            frmServer.Socket(I).Close
            frmServer.Socket(I).Accept SocketId
            Call SocketConnected(I)
        End If
    End If

End Sub

Sub SocketConnected(ByVal index As Long)
Dim I As Long
Dim var1 As String

    If index <> 0 Then
        ' make sure they're not banned
        If Not IsBanned(GetPlayerIP(index)) Then
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M6")
            Call TextAdd(var1 & GetPlayerIP(index) & ".") '[SERVIDOR]M6
        Else
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "A3")
            Call AlertMsg(index, var1 & Options.Game_Name & ".") '[KICKBANADMIN]A3
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For I = MAX_PLAYERS To 1 Step -1
            If IsConnected(I) Then
                Player_HighIndex = I
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
        SendGUIBars index
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 25 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.Length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.Length - 4
        If pLength <= TempPlayer(index).buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).buffer.Length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long)
Dim var1 As String

    If index > 0 Then
        Call LeftGame(index)
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M5")
        Call TextAdd(var1 & GetPlayerIP(index) & ".") '[SERVIDOR]M5
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long
    Dim I As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong mapnum
    buffer.WriteString Trim$(Map(mapnum).Name)
    buffer.WriteString Trim$(Map(mapnum).Music)
    buffer.WriteString Trim$(Map(mapnum).BGS)
    buffer.WriteLong Map(mapnum).Revision
    buffer.WriteByte Map(mapnum).Moral
    buffer.WriteLong Map(mapnum).Up
    buffer.WriteLong Map(mapnum).Down
    buffer.WriteLong Map(mapnum).Left
    buffer.WriteLong Map(mapnum).Right
    buffer.WriteLong Map(mapnum).BootMap
    buffer.WriteByte Map(mapnum).BootX
    buffer.WriteByte Map(mapnum).BootY
    
    buffer.WriteLong Map(mapnum).Weather
    buffer.WriteLong Map(mapnum).WeatherIntensity
    
    buffer.WriteLong Map(mapnum).Fog
    buffer.WriteLong Map(mapnum).FogSpeed
    buffer.WriteLong Map(mapnum).FogOpacity
    
    buffer.WriteLong Map(mapnum).Red
    buffer.WriteLong Map(mapnum).Green
    buffer.WriteLong Map(mapnum).Blue
    buffer.WriteLong Map(mapnum).Alpha
    
    buffer.WriteByte Map(mapnum).MaxX
    buffer.WriteByte Map(mapnum).MaxY
    
    buffer.WriteByte Map(mapnum).DropItemsOnDeath

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY

            With Map(mapnum).Tile(x, y)
                For I = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(I).x
                    buffer.WriteLong .Layer(I).y
                    buffer.WriteLong .Layer(I).Tileset
                Next
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
                buffer.WriteByte .Cubo
                buffer.WriteLong .HP
                buffer.WriteString .Mensaje
                buffer.WriteLong .Animacion
                buffer.WriteLong .Banco
                buffer.WriteLong .Evento
                buffer.WriteLong .BancoLlave
                buffer.WriteLong .Script
                buffer.WriteLong .Timer
                buffer.WriteByte .ParteCubo
                buffer.WriteLong .SFX1
                buffer.WriteLong .SFX2
                buffer.WriteLong .Objeto
                
    
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(mapnum).NPC(x)
        buffer.WriteLong Map(mapnum).NpcSpawnType(x)
    Next

    MapCache(mapnum).Data = buffer.ToArray()
    
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim N As Long
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If I <> index Then
                s = s & GetPlayerName(I) & ", "
                N = N + 1
            End If
        End If

    Next

    If N = 0 Then
        s = "No hay otros jugadores en línea." ' [SERVER]M7
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "" & N & " jugadores conectados: " & s & "." ' '[SERVER]M8
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim buffer As clsBuffer, I As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
'========================================================================================'
'===This fixes the friends lists when blank spots occur by erasing all of the friends.==='
'===I only have this here because I used it a lot in the debugging phase of the system==='
'========================================================================================'

    'For I = 1 To Player(Index).Friends.Count
    '    SetPlayerFriendName Index, I
    'Next
    'SetPlayerFriends Index, 0
    'SavePlayer Index
    
'========================================================================================'
'========================================================================================'
    
    CheckSkills index
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerLevel(index)
    buffer.WriteLong GetPlayerPOINTS(index)
    buffer.WriteLong GetPlayerSprite(index)
    buffer.WriteLong GetPlayerMap(index)
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteLong GetPlayerClass(index)
    buffer.WriteLong GetPlayerVisible(index)
    buffer.WriteLong Abs(Player(index).WalkThrough)
    buffer.WriteLong Player(index).Follower
    
    buffer.WriteLong MAX_SKILLS
    
  

    
    'For I = 1 To MAX_ELEMENT_ATTACK
    '    buffer.WriteLong Player(index).Elements(I).Elmt_Attack
    'Next I
    
    'For I = 1 To MAX_ELEMENT_DEFENCE
    '    buffer.WriteLong Player(index).Elements(I).Elmt_Defence
    'Next I
    
    For I = 1 To MAX_SKILLS
        buffer.WriteString Skill(I).Name
        buffer.WriteLong Skill(I).MaxLvl
        buffer.WriteLong Player(index).Skills(I).Level
        buffer.WriteLong Player(index).Skills(I).EXP
        buffer.WriteLong Player(index).Skills(I).EXP_Needed
    Next I
    
    For I = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, I)
    Next
    
    For I = 1 To MAX_COMBAT
        buffer.WriteByte GetPlayerCombatLevel(index, I)
        buffer.WriteLong GetPlayerCombatExp(index, I)
        buffer.WriteLong GetPlayerNextCombatLevel(index, I)
    Next
    
    If Player(index).GuildFileId > 0 Then
        buffer.WriteByte 1
        buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name
        buffer.WriteString GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag
        buffer.WriteInteger GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color
    Else
        buffer.WriteByte 0
    End If
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Function PlayerFriends(ByVal index As Long) As Byte()
Dim buffer As clsBuffer, I As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    Set buffer = New clsBuffer
    
    buffer.WriteLong SFriends
    buffer.WriteLong GetPlayerFriends(index)
    If GetPlayerFriends(index) > 0 Then
        If GetPlayerFriends(index) > MAX_FRIENDS Then Call SetPlayerFriends(index, MAX_FRIENDS, False)
        For I = 1 To GetPlayerFriends(index)
            buffer.WriteString GetPlayerFriendName(index, I)
        Next I
    End If
    
    PlayerFriends = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            If I <> index Then
                If GetPlayerMap(I) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(I)
                End If
            End If
        End If
    Next

    ' Send index's player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    SendUpdateFriendsLists
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToMapBut index, mapnum, buffer.ToArray()
    
    If TempPlayer(index).inParty > 0 Then
        SendPartyVitals TempPlayer(index).inParty, index
    End If
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendOpenBook(ByVal index As Long, ByVal libroslot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenBook
    buffer.WriteLong libroslot
    
    Call SendDataTo(index, buffer.ToArray())
    
    Set buffer = Nothing
End Sub

Sub SendPlayerFriends(ByVal index As Long)
    Dim packet As String
    SendDataToMap GetPlayerMap(index), PlayerFriends(index)
End Sub

Sub AskForFriendshipFrom(ByVal index As Long, ByVal FromWho As String)
    Dim buffer As clsBuffer
    
    If Not Len(FromWho) > 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFriendRequest
    buffer.WriteString FromWho
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub

Sub SendMap(ByVal index As Long, ByVal mapnum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(mapnum).Data) - LBound(MapCache(mapnum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(mapnum).Data()
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapnum, I).playerName
        buffer.WriteLong MapItem(mapnum, I).num
        buffer.WriteLong MapItem(mapnum, I).Value
        buffer.WriteLong MapItem(mapnum, I).x
        buffer.WriteLong MapItem(mapnum, I).y
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsToAll(ByVal mapnum As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData

    For I = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(mapnum, I).playerName
        buffer.WriteLong MapItem(mapnum, I).num
        buffer.WriteLong MapItem(mapnum, I).Value
        buffer.WriteLong MapItem(mapnum, I).x
        buffer.WriteLong MapItem(mapnum, I).y
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal mapnum As Long, ByVal mapNpcNum As Long)
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcVitals
    buffer.WriteLong mapNpcNum
    For I = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(mapnum).NPC(mapNpcNum).Vital(I)
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(I).num
        buffer.WriteLong MapNpc(mapnum).NPC(I).x
        buffer.WriteLong MapNpc(mapnum).NPC(I).y
        buffer.WriteLong MapNpc(mapnum).NPC(I).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(I).Vital(HP)
        buffer.WriteLong MapNpc(mapnum).NPC(I).HPSetTo
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNpcData

    For I = 1 To MAX_MAP_NPCS
        buffer.WriteLong MapNpc(mapnum).NPC(I).num
        buffer.WriteLong MapNpc(mapnum).NPC(I).x
        buffer.WriteLong MapNpc(mapnum).NPC(I).y
        buffer.WriteLong MapNpc(mapnum).NPC(I).Dir
        buffer.WriteLong MapNpc(mapnum).NPC(I).Vital(HP)
        buffer.WriteLong MapNpc(mapnum).NPC(I).HPSetTo
    Next

    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendItems(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_ITEMS

        If LenB(Trim$(Item(I).Name)) > 0 Then
            Call SendUpdateItemTo(index, I)
        End If

    Next

End Sub

Sub SendCombos(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_COMBOS

        If LenB(Trim$(Combo(I).Item_1)) > 0 Then
            Call SendUpdateComboTo(index, I)
        End If

    Next

End Sub

Sub SendSkills(ByVal index As Long)
    Dim I As Long
    
    For I = 1 To MAX_SKILLS
    
        If LenB(Trim$(Skill(I).Name)) > 0 Then
            Call SendUpdateSkillTo(index, I)
        End If
    
    Next
End Sub

Sub SendAnimations(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS

        If LenB(Trim$(Animation(I).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, I)
        End If

    Next

End Sub

Sub SendNpcs(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_NPCS

        If LenB(Trim$(NPC(I).Name)) > 0 Then
            Call SendUpdateNpcTo(index, I)
        End If

    Next

End Sub

Sub SendResources(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_RESOURCES

        If LenB(Trim$(Resource(I).Name)) > 0 Then
            Call SendUpdateResourceTo(index, I)
        End If

    Next

End Sub

Sub SendInventory(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For I = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(index, I)
        buffer.WriteLong GetPlayerInvItemValue(index, I)
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteLong invSlot
    buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerWornEq
    buffer.WriteLong GetPlayerEquipment(index, Armor)
    buffer.WriteLong GetPlayerEquipment(index, Weapon)
    buffer.WriteLong GetPlayerEquipment(index, Helmet)
    buffer.WriteLong GetPlayerEquipment(index, Legs) ' New
buffer.WriteLong GetPlayerEquipment(index, Boots) ' New
buffer.WriteLong GetPlayerEquipment(index, Glove) ' New
buffer.WriteLong GetPlayerEquipment(index, Ring) ' New
buffer.WriteLong GetPlayerEquipment(index, Enchant) ' New
    buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong index
    buffer.WriteLong GetPlayerEquipment(index, Armor)
    buffer.WriteLong GetPlayerEquipment(index, Weapon)
    buffer.WriteLong GetPlayerEquipment(index, Helmet)
    buffer.WriteLong GetPlayerEquipment(index, Legs)
buffer.WriteLong GetPlayerEquipment(index, Boots)
buffer.WriteLong GetPlayerEquipment(index, Glove)
buffer.WriteLong GetPlayerEquipment(index, Ring)
buffer.WriteLong GetPlayerEquipment(index, Enchant)
    buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong PlayerNum
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Armor)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Weapon)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Helmet)
    buffer.WriteLong GetPlayerEquipment(index, Legs)
buffer.WriteLong GetPlayerEquipment(index, Boots)
buffer.WriteLong GetPlayerEquipment(index, Glove)
buffer.WriteLong GetPlayerEquipment(index, Ring)
buffer.WriteLong GetPlayerEquipment(index, Enchant)
    buffer.WriteLong GetPlayerEquipment(PlayerNum, Shield)
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHp
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMp
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendEXP(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCombatEXP(ByVal index As Long)
Dim buffer As clsBuffer
Dim I As Byte

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerCombatEXP
    For I = 1 To MAX_COMBAT
        buffer.WriteByte GetPlayerCombatLevel(index, I)
        buffer.WriteLong GetPlayerCombatExp(index, I)
        buffer.WriteLong GetPlayerNextCombatLevel(index, I)
    Next
        
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGUIBars(ByVal index As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SGUIBars
    buffer.WriteLong Options.OriginalGUIBars
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGUIBarsToAll()
Dim I As Long
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            If IsConnected(I) Then
                SendGUIBars I
            End If
        End If
    Next I
End Sub

Sub SendStats(ByVal index As Long)
Dim I As Long
Dim packet As String
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    For I = 1 To Stats.Stat_Count - 1
        buffer.WriteLong GetPlayerStat(index, I)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHighlight(ByVal index As Long, ByVal invNum As Long)
Dim buffer As New clsBuffer
    
    Set buffer = New clsBuffer
    
    ' send the highlight to the client
    buffer.WriteLong SHighlightItem
    buffer.WriteLong invNum
    buffer.WriteLong Player(index).Inv(invNum).Selected
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If
    
    ' Send visibility message
    If GetPlayerAccess(index) > ADMIN_MONITOR Then
        If GetPlayerVisible(index) = 1 Then
            Call PlayerMsg(index, "(invisible)", AlertColor)
        End If
    End If

    ' Send whos online
    Call SendWhosOnline(index)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim packet As String
    Dim I As Long, N As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SClassesData
    buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        buffer.WriteString GetClassName(I)
        buffer.WriteLong GetClassMaxVital(I, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(I, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(I).MaleSprite)
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(I).FemaleSprite)
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).FemaleSprite(q)
        Next
        
        N = UBound(Class(I).Oculto)
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).Oculto(q)
        Next
        
        
        N = UBound(Class(I).VCaminar)
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).VCaminar(q)
        Next
        
        N = UBound(Class(I).VCorrer)
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).VCorrer(q)
        Next
        
                N = UBound(Class(I).Faccion) 'EaSee 0.6
        
        ' send array size
        buffer.WriteLong N
        
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).Faccion(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(I).stat(q)
        Next
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim packet As String
    Dim I As Long, N As Long, q As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    buffer.WriteLong Max_Classes

    For I = 1 To Max_Classes
        buffer.WriteString GetClassName(I)
        buffer.WriteLong GetClassMaxVital(I, Vitals.HP)
        buffer.WriteLong GetClassMaxVital(I, Vitals.MP)
        
        ' set sprite array size
        N = UBound(Class(I).MaleSprite)
        ' send array size
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).MaleSprite(q)
        Next
        
        ' set sprite array size
        N = UBound(Class(I).FemaleSprite)
        ' send array size
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).FemaleSprite(q)
        Next
        
        N = UBound(Class(I).Oculto)
        
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).Oculto(q)
        Next
        
        N = UBound(Class(I).VCaminar)
        
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).VCaminar(q)
        Next
        
        N = UBound(Class(I).VCorrer)
        
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).VCorrer(q)
        Next
        
                N = UBound(Class(I).Faccion)
        
        buffer.WriteLong N
        ' loop around sending each sprite
        For q = 0 To N
            buffer.WriteLong Class(I).Faccion(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            buffer.WriteLong Class(I).stat(q)
        Next
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString vbNullString
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    buffer.WriteLong 0
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYMap
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateComboToAll(ByVal comboNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ComboSize As Long
    Dim ComboData() As Byte
    Set buffer = New clsBuffer
    ComboSize = LenB(Combo(comboNum))
    
    ReDim ComboData(ComboSize - 1)
    
    CopyMemory ComboData(0), ByVal VarPtr(Combo(comboNum)), ComboSize
    
    buffer.WriteLong SUpdateCombo
    buffer.WriteLong comboNum
    buffer.WriteBytes ComboData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSkillToAll(ByVal index As Long, ByVal tSkill As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SkillSize As Long
    Dim SkillData() As Byte
    Set buffer = New clsBuffer
    SkillSize = LenB(Skill(tSkill))
    
    ReDim SkillData(SkillSize - 1)
    
    CopyMemory SkillData(0), ByVal VarPtr(Skill(tSkill)), SkillSize
    
    buffer.WriteLong SUpdateSkill
    buffer.WriteLong tSkill
    buffer.WriteBytes SkillData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal itemnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemnum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemnum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong itemnum
    buffer.WriteBytes ItemData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateComboTo(ByVal index As Long, ByVal comboNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ComboSize As Long
    Dim ComboData() As Byte
    Set buffer = New clsBuffer
    ComboSize = LenB(Combo(comboNum))
    ReDim ComboData(ComboSize - 1)
    CopyMemory ComboData(0), ByVal VarPtr(Combo(comboNum)), ComboSize
    buffer.WriteLong SUpdateCombo
    buffer.WriteLong comboNum
    buffer.WriteBytes ComboData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSkillTo(ByVal index As Long, ByVal tSkill As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SkillSize As Long
    Dim SData() As Byte
    Set buffer = New clsBuffer
    SkillSize = LenB(Skill(tSkill))
    ReDim SData(SkillSize - 1)
    CopyMemory SData(0), ByVal VarPtr(Skill(tSkill)), SkillSize
    buffer.WriteLong SUpdateSkill
    buffer.WriteLong tSkill
    buffer.WriteBytes SData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    Set buffer = New clsBuffer
    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(npcNum)), NPCSize
    buffer.WriteLong SUpdateNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NPCData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(I).Name)) > 0 Then
            Call SendUpdateShopTo(index, I)
        End If

    Next

End Sub

Sub SendUpdateShopToAll(ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal shopNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong shopNum
    buffer.WriteBytes ShopData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
Dim I As Long

For I = 1 To MAX_SPELLS

If LenB(Trim$(Spell(I).Name)) > 0 Then
Call SendSpellTo(index, I)
End If

Next
Call SendPlayerSpells(index)

End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim packet As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpells

    For I = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(index, I)
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim I As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then

        For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(I).ResourceState
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(I).x
            buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(I).y
        Next

    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal mapnum As Long, ByVal Resource_num As Long)
    Dim buffer As clsBuffer
    Dim I As Long
    Set buffer = New clsBuffer
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(mapnum).Resource_Count

    If ResourceCache(mapnum).Resource_Count > 0 Then

        For I = 0 To ResourceCache(mapnum).Resource_Count
            buffer.WriteByte ResourceCache(mapnum).ResourceData(I).ResourceState
            buffer.WriteLong ResourceCache(mapnum).ResourceData(I).x
            buffer.WriteLong ResourceCache(mapnum).ResourceData(I).y
        Next

    End If

    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal mapnum As Long, ByVal message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong x
    buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap mapnum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong x
    buffer.WriteLong y
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendAnimation(ByVal mapnum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, buffer.ToArray
    Else
        SendDataToMap mapnum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCooldown
    buffer.WriteLong Slot
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal mapnum As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Map] "
    buffer.WriteLong saycolour
    
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[Global] "
    buffer.WriteLong saycolour
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBank
    
    For I = 1 To MAX_BANK
        buffer.WriteLong Bank(index).Item(I).num
        buffer.WriteLong Bank(index).Item(I).Value
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal mapnum As Long, ByVal x As Long, ByVal y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteByte Value
    SendDataToMap mapnum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal shopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong shopNum
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteLong GetPlayerDir(index)
    buffer.WriteLong movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    buffer.WriteLong tradeTarget
    buffer.WriteString Trim$(GetPlayerName(tradeTarget))
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal dataType As Byte)
Dim buffer As clsBuffer
Dim I As Long
Dim tradeTarget As Long
Dim totalWorth As Long
    
    tradeTarget = TempPlayer(index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    buffer.WriteByte dataType
    
    If dataType = 0 Then ' own inventory
        For I = 1 To MAX_INV
            buffer.WriteLong TempPlayer(index).TradeOffer(I).num
            buffer.WriteLong TempPlayer(index).TradeOffer(I).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(I).num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(I).num).Type = ITEM_TYPE_CURRENCY Or Item(TempPlayer(index).TradeOffer(I).num).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(I).num)).price * TempPlayer(index).TradeOffer(I).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(I).num)).price
                End If
            End If
        Next
    ElseIf dataType = 1 Then ' other inventory
        For I = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)
            buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(I).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)).Stackable > 0 Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)).price * TempPlayer(tradeTarget).TradeOffer(I).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)).price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    buffer.WriteLong totalWorth
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    buffer.WriteByte Status
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendTarget(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(index).target
    buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
Dim I As Long
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    For I = 1 To MAX_HOTBAR
        buffer.WriteLong Player(index).Hotbar(I).Slot
        buffer.WriteByte Player(index).Hotbar(I).sType
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLoginOk(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    buffer.WriteLong index
    buffer.WriteLong Player_HighIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHighIndex()
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong Player_HighIndex
    buffer.WriteLong Options.FullScreen
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    buffer.WriteLong x
    buffer.WriteLong y
    buffer.WriteLong entityType
    buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteString Trim$(Player(TradeRequest).Name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal targetPlayer As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    buffer.WriteString Trim$(Player(targetPlayer).Name)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
Dim buffer As clsBuffer, I As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    buffer.WriteByte 1
    buffer.WriteLong Party(partyNum).Leader
    For I = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(partyNum).Member(I)
    Next
    buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
Dim buffer As clsBuffer, I As Long, partyNum As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' check if we're in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(partyNum).Leader
        For I = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(partyNum).Member(I)
        Next
        buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
Dim buffer As clsBuffer, I As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    buffer.WriteLong index
    For I = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(index, I)
        buffer.WriteLong Player(index).Vital(I)
    Next
        buffer.WriteString Player(index).Name
    SendDataToParty partyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpawnItemToMap(ByVal mapnum As Long, ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnItem
    buffer.WriteLong index
    buffer.WriteString MapItem(mapnum, index).playerName
    buffer.WriteLong MapItem(mapnum, index).num
    buffer.WriteLong MapItem(mapnum, index).Value
    buffer.WriteLong MapItem(mapnum, index).x
    buffer.WriteLong MapItem(mapnum, index).y
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChatBubble(ByVal mapnum As Long, ByVal target As Long, ByVal targetType As Long, ByVal message As String, ByVal Colour As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    buffer.WriteLong target
    buffer.WriteLong targetType
    buffer.WriteString message
    buffer.WriteLong Colour
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'fognum
            buffer.WriteLong Data2 'fog movement speed
            buffer.WriteLong Data3 'opacity
        Case EFFECT_TYPE_WEATHER
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'weather type
            buffer.WriteLong Data2 'weather intensity
        Case EFFECT_TYPE_TINT
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 'red
            buffer.WriteLong Data2 'green
            buffer.WriteLong Data3 'blue
            buffer.WriteLong Data4 'alpha
    End Select
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub


Sub SendAttack(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong ServerPackets.SAttack
    buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendFlash(ByVal target As Long, mapnum As Long, isNpc As Boolean)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFlash
    buffer.WriteLong target
    If isNpc Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendProjectileToMap(ByVal index As Long, ByVal PlayerProjectile As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SHandleProjectile
    buffer.WriteLong PlayerProjectile
    buffer.WriteLong index
    With TempPlayer(index).ProjecTile(PlayerProjectile)
        buffer.WriteLong .Direction
        buffer.WriteLong .Pic
        buffer.WriteLong .Range
        buffer.WriteLong .Damage
        buffer.WriteLong .Speed
        buffer.WriteLong .Municion
    End With
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVelocidad(ByVal caminar As Long, ByVal correr As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SVelocidad
    buffer.WriteLong caminar
    buffer.WriteByte correr
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
    
    Exit Sub

End Sub

Sub SendClima(ByVal mapa As Long, ByVal Clima As Byte, ByVal intensidad As Byte, ByVal niebla As Byte, ByVal Velocidad As Byte, ByVal opacidad As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClima
    buffer.WriteLong mapa
    buffer.WriteByte Clima
    buffer.WriteByte intensidad
    buffer.WriteByte niebla
    buffer.WriteByte Velocidad
    buffer.WriteByte opacidad
    
    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
    
    Exit Sub
End Sub

Sub EnviarLetrero(ByVal index As Long, ByVal texto As String)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SLetrero
    buffer.WriteString texto
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
    
    Exit Sub

End Sub


Sub SendClimaBas(ByVal Clima As Byte)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SClimabas
    buffer.WriteByte Clima

    SendDataToAll buffer.ToArray()
    
    Set buffer = Nothing
    
    Exit Sub
End Sub

Sub EnviarConfusion(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarConfusion
    buffer.WriteLong TempPlayer(index).ConfusionDuracion
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub EnviarParalisis(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarParalisis
    buffer.WriteLong TempPlayer(index).ParalisisDuracion
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub EnviarVeneno(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarVeneno
    buffer.WriteLong TempPlayer(index).VenenoDuracion
    buffer.WriteLong TempPlayer(index).VenenoGolpe
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub


Sub EnviarVelocidad(ByVal index As Long) 'EaSee 0.5
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarVelocidad
    buffer.WriteLong TempPlayer(index).VelocidadDuracion
    buffer.WriteLong TempPlayer(index).VelocidadCaminar2
    buffer.WriteLong TempPlayer(index).VelocidadCorrer2
    buffer.WriteLong TempPlayer(index).VelocidadBuff

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub EnviarInvisibilidad(ByVal index As Long) 'EaSee 0.5

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarInvisibilidadH
    buffer.WriteLong TempPlayer(index).InvisibilidadDuracion
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

    Player(index).Visible = 1
    
    Call SendPlayerData(index)

End Sub

Sub EnviarFuerzaH(ByVal index As Long) 'EaSee 0.5
Dim StatTemporal As Long

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarFuerzaH
    buffer.WriteLong TempPlayer(index).FuerzaDuracion
    buffer.WriteLong TempPlayer(index).BuffValor
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    If TempPlayer(index).BuffValor = 1 Then 'Positivo
    Player(index).FuerzaBuff = Player(index).FuerzaBuff + TempPlayer(index).FuerzaH
    StatTemporal = Stats.strength - Player(index).FuerzaBuff 'Restauramos a original si ya hay activo para sumarlos
    Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) + TempPlayer(index).FuerzaH)
    Else 'Negativo
    If TempPlayer(index).FuerzaH > Stats.strength Then Exit Sub
    Player(index).FuerzaDeBuff = Player(index).FuerzaDeBuff + TempPlayer(index).FuerzaH
    StatTemporal = Stats.strength + Player(index).FuerzaDeBuff 'Original
    Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) - TempPlayer(index).FuerzaH)
    End If
    
    Call SendPlayerData(index)

End Sub

Sub EnviarDestrezaH(ByVal index As Long) 'EaSee 0.5
Dim StatTemporal As Long

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarDestrezaH
    buffer.WriteLong TempPlayer(index).DestrezaDuracion
    buffer.WriteLong TempPlayer(index).BuffValor
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    If TempPlayer(index).BuffValor = 1 Then 'Positivo
    Player(index).DestrezaBuff = Player(index).DestrezaBuff + TempPlayer(index).DestrezaH
    StatTemporal = Stats.Endurance - Player(index).DestrezaBuff 'Restauramos a original si ya hay activo para sumarlos
    Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + TempPlayer(index).DestrezaH)
    Else 'Negativo
        If TempPlayer(index).DestrezaH > Stats.Endurance Then Exit Sub
    Player(index).DestrezaDeBuff = Player(index).DestrezaDeBuff + TempPlayer(index).DestrezaH
    StatTemporal = Stats.Endurance + Player(index).DestrezaDeBuff 'Original
    Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) - TempPlayer(index).DestrezaH)
    End If
    
    Call SendPlayerData(index)

End Sub


Sub EnviarAgilidadH(ByVal index As Long) 'EaSee 0.5
Dim StatTemporal As Long

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarAgilidadH
    buffer.WriteLong TempPlayer(index).AgilidadDuracion
    buffer.WriteLong TempPlayer(index).BuffValor
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    If TempPlayer(index).BuffValor = 1 Then 'Positivo
    Player(index).AgilidadBuff = Player(index).AgilidadBuff + TempPlayer(index).AgilidadH
    StatTemporal = Stats.Agility - Player(index).AgilidadBuff 'Restauramos a original si ya hay activo para sumarlos
    Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + TempPlayer(index).AgilidadH)
    Else 'Negativo
        If TempPlayer(index).AgilidadH > Stats.Agility Then Exit Sub
    Player(index).AgilidadDeBuff = Player(index).AgilidadDeBuff + TempPlayer(index).AgilidadH
    StatTemporal = Stats.Agility + Player(index).AgilidadDeBuff 'Original
    Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) - TempPlayer(index).AgilidadH)
    End If
    
    Call SendPlayerData(index)

End Sub

Sub EnviarInteligenciaH(ByVal index As Long) 'EaSee 0.5
Dim StatTemporal As Long

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarInteligenciaH
    buffer.WriteLong TempPlayer(index).InteligenciaDuracion
    buffer.WriteLong TempPlayer(index).BuffValor
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    If TempPlayer(index).BuffValor = 1 Then 'Positivo
    Player(index).InteligenciaBuff = Player(index).InteligenciaBuff + TempPlayer(index).InteligenciaH
    StatTemporal = Stats.Intelligence - Player(index).InteligenciaBuff 'Restauramos a original si ya hay activo para sumarlos
    Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + TempPlayer(index).InteligenciaH)
    Else 'Negativo
        If TempPlayer(index).InteligenciaH > Stats.Intelligence Then Exit Sub
        Player(index).InteligenciaDeBuff = Player(index).InteligenciaDeBuff + TempPlayer(index).InteligenciaH
    StatTemporal = Stats.Intelligence + Player(index).InteligenciaDeBuff 'Original
    Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) - TempPlayer(index).InteligenciaH)
    End If
    
    Call SendPlayerData(index)

End Sub

Sub EnviarVoluntadH(ByVal index As Long) 'EaSee 0.5
Dim StatTemporal As Long

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarVoluntadH
    buffer.WriteLong TempPlayer(index).VoluntadDuracion
    buffer.WriteLong TempPlayer(index).BuffValor
 
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    If TempPlayer(index).BuffValor = 1 Then 'Positivo
    Player(index).VoluntadBuff = Player(index).VoluntadBuff + TempPlayer(index).VoluntadH
    StatTemporal = Stats.Willpower - Player(index).VoluntadBuff 'Restauramos a original si ya hay activo para sumarlos
    Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + TempPlayer(index).VoluntadH)
    Else 'Negativo
   If TempPlayer(index).VoluntadH > Stats.Willpower Then Exit Sub
   Player(index).VoluntadDeBuff = Player(index).VoluntadDeBuff + TempPlayer(index).VoluntadH
    StatTemporal = Stats.Willpower + Player(index).VoluntadDeBuff 'Original
    Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) - TempPlayer(index).VoluntadH)
    End If
    
    Call SendPlayerData(index)

End Sub

Sub EnviarSprite(ByVal index As Long) 'EaSee 0.5

 Player(index).Sprite = TempPlayer(index).SpriteNumero
 SendPlayerData (index)

 Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnviarSpriteH
    buffer.WriteLong TempPlayer(index).SpriteDuracion
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing


End Sub

Sub EnviarTransportarH(ByVal index As Long) 'EaSee 0.5 '[SERVIDOR]M9 y M10
Dim var1 As String

                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M9")
    Call PlayerWarp(index, TempPlayer(index).TransMapaH, TempPlayer(index).TransMapaxH, TempPlayer(index).TransMapayH)
    Call PlayerMsg(index, var1 & TempPlayer(index).TransMapaH, BrightBlue)
                        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "SERVIDOR", "M10")
    Call AddLog(GetPlayerName(index) & var1 & TempPlayer(index).TransMapaH & ".", ADMIN_LOG)
End Sub

Sub SendSpellTo(ByVal index As Long, ByVal spellnum As Long)
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte

Set buffer = New clsBuffer

SpellSize = LenB(Spell(spellnum))
ReDim SpellData(SpellSize - 1)
CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize

buffer.WriteLong SSpell
buffer.WriteLong spellnum
buffer.WriteBytes SpellData

SendDataTo index, buffer.ToArray()
Set buffer = Nothing
End Sub

Sub EnviarSpriteColorH(ByVal index As Long, ByVal tipo As Long, ByVal Duracion As Long) 'EaSee 0.6
 Dim buffer As clsBuffer 'EaSee 0.6
    
    Set buffer = New clsBuffer
    buffer.WriteLong SColorSpriteH
    buffer.WriteLong index 'Jugador Afectado
    buffer.WriteLong tipo 'Hechizo
    buffer.WriteLong Duracion 'Tiempo
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing


End Sub

Sub SendSpriteAnimAtaq(ByVal index As Long)
Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimAtaque
    buffer.WriteLong Options.AnimacionAtaque
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpriteAnimAtaqToAll()
Dim I As Long
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            If IsConnected(I) Then
                SendSpriteAnimAtaq I
            End If
        End If
    Next I
End Sub

