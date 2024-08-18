Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CBroadcastMsg) = GetAddress(AddressOf HandleBroadcastMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CHighlightItem) = GetAddress(AddressOf HandleHighlightItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CPlayerInfoRequest) = GetAddress(AddressOf HandlePlayerInfoRequest)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CGetStats) = GetAddress(AddressOf HandleGetStats)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CBanList) = GetAddress(AddressOf HandleBanList)
    HandleDataSub(CBanDestroy) = GetAddress(AddressOf HandleBanDestroy)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNpc) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNpc) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditspell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CSetName) = GetAddress(AddressOf HandleSetName)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMotd) = GetAddress(AddressOf HandleSetMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCast) = GetAddress(AddressOf HandleCast)
    HandleDataSub(CQuit) = GetAddress(AddressOf HandleQuit)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCS) = GetAddress(AddressOf HandleRequestNPCS)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CChangeBankSlots) = GetAddress(AddressOf HandleChangeBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CHotbarUse) = GetAddress(AddressOf HandleHotbarUse)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CProjecTileAttack) = GetAddress(AddressOf HandleProjecTileAttack)
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    HandleDataSub(CPlayerVisibility) = GetAddress(AddressOf HandlePlayerVisibility)
    HandleDataSub(CHealPlayer) = GetAddress(AddressOf HandleHealPlayer)
    HandleDataSub(CKillPlayer) = GetAddress(AddressOf HandleKillPlayer)
    HandleDataSub(CSayGuild) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CGuildCommand) = GetAddress(AddressOf HandleGuildCommands)
    HandleDataSub(CSaveGuild) = GetAddress(AddressOf HandleGuildSave)
    HandleDataSub(CCharEditorCommand) = GetAddress(AddressOf HandleCharEditorCommand)
    HandleDataSub(CRequestEditQuest) = GetAddress(AddressOf HandleRequestEditQuest)
    HandleDataSub(CSaveQuest) = GetAddress(AddressOf HandleSaveQuest)
    HandleDataSub(CRequestQuests) = GetAddress(AddressOf HandleRequestQuests)
    HandleDataSub(CPlayerHandleQuest) = GetAddress(AddressOf HandlePlayerHandleQuest)
    HandleDataSub(CQuestLogUpdate) = GetAddress(AddressOf HandleQuestLogUpdate)
    HandleDataSub(COpenMyBank) = GetAddress(AddressOf HandleOpenMyBank)
    HandleDataSub(CWalkthrough) = GetAddress(AddressOf HandleToggleWalkthrough)
    HandleDataSub(CFollowPlayer) = GetAddress(AddressOf HandleStartFollowingPlayer)
    HandleDataSub(CClickPos) = GetAddress(AddressOf HandleBeFriend)
    HandleDataSub(CDeleteFriend) = GetAddress(AddressOf HandleDeleteFriend)
    HandleDataSub(CUpdateFList) = GetAddress(AddressOf HandleUpdateFriendsList)
    HandleDataSub(CFriendAccept) = GetAddress(AddressOf HandleAcceptFriend)
    HandleDataSub(CFriendDecline) = GetAddress(AddressOf HandleDeclineFriend)
    HandleDataSub(CPrivateMsg) = GetAddress(AddressOf HandlePrivateMsg)
    HandleDataSub(CRequestFriendData) = GetAddress(AddressOf HandleRequestFriendData)
    HandleDataSub(CRequestEditCombo) = GetAddress(AddressOf HandleRequestEditCombos)
    HandleDataSub(CRequestCombos) = GetAddress(AddressOf HandleRequestCombos)
    HandleDataSub(CSaveCombo) = GetAddress(AddressOf HandleSaveCombo)
    HandleDataSub(CInvHidden) = GetAddress(AddressOf HandleInvHidden)
    HandleDataSub(CEnviarMapaCubos) = GetAddress(AddressOf ProcesarCuboMapa)
    HandleDataSub(CEnviarMapaCubos64) = GetAddress(AddressOf ProcesarCuboMapa64) 'by ikicapitan
    HandleDataSub(CMapaRespawnCubos) = GetAddress(AddressOf MapaRespawnCubos)
    HandleDataSub(CGolpe) = GetAddress(AddressOf Procesar_Golpe)
    HandleDataSub(CEstadosBuff) = GetAddress(AddressOf Procesar_Buff)
    HandleDataSub(CEnviarVisibilidad) = GetAddress(AddressOf Procesar_Visibilidad) 'EaSee 0.5
    HandleDataSub(CRestaurarSprite) = GetAddress(AddressOf Procesar_RestaurarSprite)
    HandleDataSub(CSendCommandForClient) = GetAddress(AddressOf HandleAcComando)
    HandleDataSub(COrderAchievements) = GetAddress(AddressOf HandleOrderAchievements)
    HandleDataSub(CAttackCubo) = GetAddress(AddressOf HandleCubo)
    
    
    
End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long
        
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then
        Exit Sub
    End If
    
    If MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Sub HandleInvHidden(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
    For I = 1 To MAX_INV
        If Player(index).Inv(I).Selected = 1 Then
            Player(index).Inv(I).Selected = 0
            SendHighlight index, I
            Exit Sub
        End If
    Next I
End Sub

Sub HandleUpdateFriendsList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call UpdateFriendsList(index)
End Sub

Sub HandleDeleteFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim fName As String, I As Long
Dim Parse() As String

    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    fName = Trim$(buffer.ReadString)
    Parse() = Split(fName, " ")
    fName = Parse(0)
    I = FindPlayer(fName)
    
    'Is there a name in the Variable?
    If Not Len(fName) > 0 Then Exit Sub
    
    ' Name's good, remove name from the list of both players
    Call RemoveFriend(index, fName)
    
    ' Tell the players
    Call PlayerMsg(index, "Han dejado de ser amigo del jugador " & fName, BrightRed) ' MateoD
    Call PlayerMsg(I, "Has dejado de ser amigo del jugador " & GetPlayerName(index), BrightRed) ' MateoD
    
    ' Send the data
    SendDataTo index, PlayerFriends(index)
    SendDataTo I, PlayerFriends(I)
    
    Set buffer = Nothing
End Sub

Sub HandleStartFollowingPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim WhoToFollow As Long
Dim I As Long
    Dim buffer As New clsBuffer
    buffer.WriteBytes Data()
    WhoToFollow = buffer.ReadLong
    
    ' Make sure we're not following anyone else
    For I = 1 To MAX_PLAYERS
        If Player(I).Follower = index Then
            Player(I).Follower = 0
            Call SendPlayerData(I)
            Exit For
        End If
    Next I
    
    If FollowerIsNearMe(index, WhoToFollow, False) Then
        Player(WhoToFollow).Follower = index
        Call PlayerMsg(index, "Estás siguiendo a " & GetPlayerName(WhoToFollow), BrightBlue) ' MateoD
    Else
        Call PlayerMsg(index, "Debes estar próximo al jugador para seguirlo.", Red) ' MateoD
    End If
    Set buffer = Nothing
End Sub

Public Sub HandleToggleWalkthrough(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim WalkThrough As Boolean
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub
    WalkThrough = Player(index).WalkThrough
    Player(index).WalkThrough = Not WalkThrough
    If WalkThrough Then Call PlayerMsg(index, "Tutorial desactivado.", White) ' MateoD
    If Not WalkThrough Then Call PlayerMsg(index, "Tutorial activado.", White) ' MateoD
    Call SendPlayerData(index)
End Sub

Private Sub HandleOpenMyBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendBank index
    TempPlayer(index).InBank = True
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim I As Long
    Dim N As Long

    If Not isPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "El usuario y contraseña deben poseer entre 3 y 12 caracteres.") ' MateoD
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(index, "El usuario y contraseña deben poseer entre 3 y 12 caracteres.") ' MateoD
                Exit Sub
            End If

            ' Prevent hacking
            For I = 1 To Len(Name)
                N = AscW(Mid$(Name, I, 1))

                If Not isNameLegal(N) Then
                    Call AlertMsg(index, "Nombre inválido, no se permiten caracteres especiales.") ' MateoD
                    Exit Sub
                End If

            Next

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd("Usuario " & Name & " creado con éxito.") ' MateoD
                Call AddLog("Usuario " & Name & " creado.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)
                
                ' Check if character data has been created
                If LenB(Trim$(Player(index).Name)) > 0 Then
                    ' we have a char!
                    HandleUseChar index
                Else
                    ' send new char shit
                    If Not isPlaying(index) Then
                        Call SendNewCharClasses(index)
                    End If
                End If
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(index) & " ha iniciado sesion desde " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " ha iniciado sesion desde " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "El usuario ya existe.") ' MateoD
            End If
            
            Set buffer = Nothing
        End If
    End If

End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim I As Long

    If Not isPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "El usuario y contraseña deben poseer entre 3 y 12 caracteres") ' MateoD
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                  Call AlertMsg(index, "El usuario no existe.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Contraseña errónea.") ' MateoD
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(Player(index).Name)) > 0 Then
                Call DeleteName(Player(index).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " ha sido eliminado.", PLAYER_LOG)
            Call AlertMsg(index, "Cuenta eliminada.")
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim I As Long
    Dim N As Long

    If Not isPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            ' Get the data
            Name = buffer.ReadString
            Password = buffer.ReadString

            ' Check versions
            If buffer.ReadLong < CLIENT_MAJOR Or buffer.ReadLong < CLIENT_MINOR Or buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Tu version del cliente esta desactualizada o no es valida con el servidor actual!")
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "The server is bein shut down.") ' MateoD
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "El usuario o la contraseña son muy cortos") ' MateoD
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                 If MDE = 1 Then
                  Call AddAccount(index, Name, Password)
                 Else
                  Call AlertMsg(index, "El usuario no existe.")
                  Exit Sub
                 End If
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Contraseña incorrecta.") ' MateoD
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
            CloseSocket (FindPlayer(Name))
            End If
            
            If MDE = 1 Then
             If TempPlayer(1).InGame = True Then
              If TempPlayer(2).InGame = True Then
               Call AlertMsg(index, "No se pueden conectar mas de 2 personas a una partida SinglePlayer.")
                CloseSocket (FindPlayer(Name))
                 Exit Sub
              End If
             End If
            End If
            
            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            ' check skill stats
            Call CheckSkills(index)
            
            ' Check if character data has been created
            If LenB(Trim$(Player(index).Name)) > 0 Then
                ' we have a char!
                HandleUseChar index
            Else
                ' send new char shit
                If Not isPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
            
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " inició sesion desde " & GetPlayerIP(index) & ".", PLAYER_LOG) ' MateoD
            Call TextAdd(GetPlayerLogin(index) & " inició sesion desde " & GetPlayerIP(index) & ".") ' MateoD
            
            Set buffer = Nothing
        End If
    End If

End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim I As Long
    Dim N As Long
Dim caminar As Long
Dim correr As Byte
Dim filename As String
filename = App.Path & "\data\classes.ini"

    If Not isPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        Name = buffer.ReadString
        Sex = buffer.ReadLong
        Class = buffer.ReadLong
        Sprite = buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "El usuario y contraseña deben poseer entre 3 y 12 caracteres.") ' MateoD
            Exit Sub
        End If

        ' Prevent hacking
        For I = 1 To Len(Name)
            N = AscW(Mid$(Name, I, 1))

            If Not isNameLegal(N) Then
                Call AlertMsg(index, "Nombre inválido, no se permiten caracteres especiales.") ' MateoD
                Exit Sub
            End If

        Next

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "Personaje ya existente.")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Este nombre ya está en uso.")
            Exit Sub
        End If


        ' Everything went ok, add the character
        

        Call AddChar(index, Name, Sex, Class, Sprite)
        Call AddLog("Personaje " & Name & " agregado a " & GetPlayerLogin(index) & " cuenta.", PLAYER_LOG)
        ' log them in!!
        HandleUseChar index
        
        
        Set buffer = Nothing
    End If

End Sub
' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, I, 1)) < 128 Or AscW(Mid$(Msg, I, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, I, 1)) < 224 Or AscW(Mid$(Msg, I, 1)) > 253 Then
                    Mid$(Msg, I, 1) = ""
                End If
            End If
        End If
    Next

    Call AddLog("Mapa #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " dice, '" & Msg & "'", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, Msg, QBColor(White))
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Msg, White)
    
    Set buffer = Nothing
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)

        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            Exit Sub
        End If

    Next

    Call AddLog("Mapa #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set buffer = Nothing
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)

        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            Exit Sub
        End If

    Next

    s = "[Global]" & GetPlayerName(index) & ": " & Msg
    Call SayMsg_Global(index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set buffer = Nothing
End Sub

Private Sub HandlePrivateMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim I As Long
    Dim MsgTo As Long, OrigName As String
    Dim Continue As Boolean
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    OrigName = buffer.ReadString
    MsgTo = FindPlayer(OrigName)
    Msg = buffer.ReadString

    ' Prevent hacking
    For I = 1 To Len(Msg)

        If AscW(Mid$(Msg, I, 1)) < 32 Or AscW(Mid$(Msg, I, 1)) > 126 Then
            Exit Sub
        End If

    Next

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        ' Make sure the two are friends.
        Continue = False
        For I = 1 To GetPlayerFriends(index)
            If LCase$(GetPlayerFriendName(index, I)) = LCase$(OrigName) Then
                Continue = True
            End If
        Next
            
        If Not Continue Then
            Call PlayerMsg(index, "Solo amigos pueden enviarse mensajes privados.", BrightRed)
            Call PlayerMsg(index, "Para enviar solicitud de amistad selecciona al objetivo y presiona la tecla B.", White) ' MateoD
            Exit Sub
        End If
            
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(index) & " dice " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
            Call PlayerMsg(MsgTo, "[PM] " & GetPlayerName(index) & ": '" & Msg & "'", TellColor)
            Call PlayerMsg(index, "[PM] " & GetPlayerName(MsgTo) & ": '" & Msg & "'", TellColor)
        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "No puedes enviarte un mensaje.", BrightRed)
    End If
    
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim movement As Long
    Dim buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    Dim I As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    movement = buffer.ReadLong 'CLng(Parse(2))
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If movement < 1 Or movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a spell
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    'Cant move if in the bank!
    If TempPlayer(index).InBank Then
        'Call SendPlayerXY(Index)
        'Exit Sub
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them from moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prevent player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If

    ' If following someone, stop
    'For I = 1 To MAX_PLAYERS
    '    If Player(Index).Follower = I Then
    '        Player(Index).Follower = 0
    '        Exit For
    '    End If
    'Next I
    Dim Path As String
            Dim caminar As Long
            Dim correr As Long
            Path = App.Path & "\data\classes.ini"
            caminar = GetVar(Path, "CLASS" & Player(index).Class, "VCaminar")
            correr = GetVar(Path, "CLASS" & Player(index).Class, "VCorrer")
            Call SendVelocidad(caminar, correr)
    Call PlayerMove(index, Dir, movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long
Dim buffer As clsBuffer
    
    ' get inventory slot number
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong
    Set buffer = Nothing

    UseItem index, invNum
    
    ' send highlight item
    Set buffer = New clsBuffer
    buffer.WriteLong SHighlightItem
    buffer.WriteLong invNum
    buffer.WriteLong Player(index).Inv(invNum).Selected
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleHighlightItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim invNum As Long, I As Long, tempNum As Long, aiiSelected As Boolean
Dim Sel1 As Boolean, Sel2 As Boolean, II As Long
Dim Sel1_Index As Long, Sel2_Index As Long
Dim reSet As Boolean
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong
    aiiSelected = False
    
    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then Exit Sub
    
        Call CheckHighlight(index, invNum)
    
    Set buffer = Nothing
    SendHighlight index, invNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim N As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index

    ' Try to attack a player
    For I = 1 To Player_HighIndex
        TempIndex = I

        ' Make sure we dont try to attack ourselves
        If TempIndex <> index Then
            TryPlayerAttackPlayer index, I
        End If
    Next

    ' Try to attack a npc
    For I = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, I
    Next

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    CheckResource index, x, y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim PointType As Byte
Dim buffer As clsBuffer
Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then
        Exit Sub
    End If

    ' Make sure they have points
    If GetPlayerPOINTS(index) > 0 Then
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= 255 Then
            PlayerMsg index, "Este estado ha alcanzado el máximo de puntos posible.", BrightRed ' MateoD
            Exit Sub
        End If
        
        ' Take away a stat point
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

        ' Everything is ok
        Select Case PointType
            Case Stats.strength
                Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) + 1)
                sMes = "Fuerza"
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
                sMes = "Resistencia"
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
                sMes = "Inteligencia"
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
                sMes = "Agilidad"
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
                sMes = "Voluntad"
        End Select
        
        SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    Else
        Exit Sub
    End If

    ' Send the update
    'Call SendStats(Index)
    SendPlayerData index
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim I As Long
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString 'Parse(1)
    Set buffer = Nothing
    I = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            Call PlayerWarp(index, GetPlayerMap(N), GetPlayerX(N), GetPlayerY(N))
            Call PlayerMsg(N, GetPlayerName(index) & " se teletransportó a ti.", BrightBlue) ' MateoD
            Call PlayerMsg(index, "Has sido teletransportado a " & GetPlayerName(N) & ".", BrightBlue) ' MateoD
            Call AddLog(GetPlayerName(index) & " ha transportado a " & GetPlayerName(N) & ", mapa #" & GetPlayerMap(N) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "Físicamente imposible", White) ' MateoD
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            Call PlayerWarp(N, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(N, "Has sido convocado por " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(N) & " ha sido convocado.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " ha transportado a " & GetPlayerName(N) & " hacia el mismo, mapa #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "Físicamente imposible.", White) ' MateoD
    End If

End Sub


' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    N = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If N < 0 Or N > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, N, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "Has sido teletransportado al mapa #" & N, BrightBlue) ' MateoD
    Call AddLog(GetPlayerName(index) & " teletransportado al mapa #" & N & ".", ADMIN_LOG) ' MateoD
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim I As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
    Exit Sub
    End If
    
    ' The sprite
    N = buffer.ReadLong 'CLng(Parse(1))
    I = FindPlayer(buffer.ReadString)
    Set buffer = Nothing
    
    Call SetPlayerSprite(I, N)
    Call SendPlayerData(I)
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Stats request packet ::
' ::::::::::::::::::::::::::
Sub HandleGetStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, Dir, 1)
End Sub
Sub HandleAcComando(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Comando As String
    Dim splitCo() As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Comando = buffer.ReadString 'CLng(Parse(1))
    Set buffer = Nothing
    ' Prevent hacking
    If CaracteristicaExist(Comando, " ") = True Then
    splitCo = Split(Comando, " ")
    'MsgBox splitCo
    Select Case UCase(splitCo(0))
    Case "/LOGROS"
    Call ReadAchievementPlayer(index, splitCo(1))
    Case "/ROOT"
        If splitCo(1) = "0000" Then Else Exit Sub
        If MDE = "1" Then
        If GetPlayerAccess(index) = 0 Then
        If isPlaying(index) And IsConnected(index) Then
        Call PlayerMsg(index, "Tienes que reiniciar el cliente para que se te den los permisos", Yellow)
        Player(index).Access = 4
        End If
        End If
        End If
    Case Else
    Call PlayerMsg(index, "Comando Desconocido: " & Comando, Red)
    End Select
    
    Else
    
    Select Case UCase(Comando)
    Case "/LOGROS"
    Call PlayerMsg(index, "/Logros <name>", Yellow)
    Case "/LOGRO"
    Call PlayerMsg(index, "/Logros <name>", Yellow)
    Case "/SERVEROK"
    Call PlayerMsg(index, "El servidor funciona perfectamente.", Yellow)
    Case Else
    If FExist(UCase(Comando)) = True Then
    Call ReadCode("" & index, Comando)
    Else
    Call PlayerMsg(index, "Comando Desconocido: " & Comando, Red)
    End If
    End Select
    
    End If

    'Call PlayerMove(index, Dir, 1)
End Sub
Sub HandleOrderAchievements(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Comando As String
    Dim max As Long
    Dim I As Long
    Dim Conjunto As String
    Dim LTiene As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Comando = buffer.ReadString 'CLng(Parse(1))
    Conjunto = ""
    If Comando = "1" Then Else Exit Sub
    
    max = GetVar(App.Path & "/data/achievement.ini", "LOGROS", "Logros")
    For I = 1 To max
    If Conjunto = "" Then
    If AchievementExist(GetVar(App.Path & "/data/achievement.ini", "PLAYERS", "" & GetPlayerName(index)), "" & I) = True Then LTiene = 1
    If AchievementExist(GetVar(App.Path & "/data/achievement.ini", "PLAYERS", "" & GetPlayerName(index)), "" & I) = False Then LTiene = 0
    Conjunto = max & "$$" & "[" & I & "]" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Nombre") & "<" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Descrip") & "<" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Imagen") & "<" & LTiene & "<Tu" & "[" & I & "]"
    Else
    If AchievementExist(GetVar(App.Path & "/data/achievement.ini", "PLAYERS", "" & GetPlayerName(index)), "" & I) = True Then LTiene = 1
    If AchievementExist(GetVar(App.Path & "/data/achievement.ini", "PLAYERS", "" & GetPlayerName(index)), "" & I) = False Then LTiene = 0
    Conjunto = Conjunto & "[" & I & "]" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Nombre") & "<" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Descrip") & "<" & GetVar(App.Path & "/data/achievement.ini", "" & I, "Imagen") & "<" & LTiene & "<Tu" & "[" & I & "]"
    End If
    Next
    
    SendAchievement index, Conjunto, False
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(index)
    I = Map(mapnum).Revision + 1
    Call ClearMap(mapnum)
    
    Map(mapnum).Name = buffer.ReadString
    Map(mapnum).Music = buffer.ReadString
    Map(mapnum).BGS = buffer.ReadString
    Map(mapnum).Revision = I
    Map(mapnum).Moral = buffer.ReadByte
    Map(mapnum).Up = buffer.ReadLong
    Map(mapnum).Down = buffer.ReadLong
    Map(mapnum).Left = buffer.ReadLong
    Map(mapnum).Right = buffer.ReadLong
    Map(mapnum).BootMap = buffer.ReadLong
    Map(mapnum).BootX = buffer.ReadByte
    Map(mapnum).BootY = buffer.ReadByte
    
    Map(mapnum).Weather = buffer.ReadLong
    Map(mapnum).WeatherIntensity = buffer.ReadLong
    
    Map(mapnum).Fog = buffer.ReadLong
    Map(mapnum).FogSpeed = buffer.ReadLong
    Map(mapnum).FogOpacity = buffer.ReadLong
    
    Map(mapnum).Red = buffer.ReadLong
    Map(mapnum).Green = buffer.ReadLong
    Map(mapnum).Blue = buffer.ReadLong
    Map(mapnum).Alpha = buffer.ReadLong
    
    Map(mapnum).MaxX = buffer.ReadByte
    Map(mapnum).MaxY = buffer.ReadByte
    
    Map(mapnum).DropItemsOnDeath = buffer.ReadByte
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Layer(I).x = buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(I).y = buffer.ReadLong
                Map(mapnum).Tile(x, y).Layer(I).Tileset = buffer.ReadLong
            Next
            For z = 1 To MapLayer.Layer_Count - 1
                Map(mapnum).Tile(x, y).Autotile(z) = buffer.ReadLong
            Next
            Map(mapnum).Tile(x, y).Type = buffer.ReadByte
            Map(mapnum).Tile(x, y).Data1 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data2 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data3 = buffer.ReadLong
            Map(mapnum).Tile(x, y).Data4 = buffer.ReadString
            Map(mapnum).Tile(x, y).DirBlock = buffer.ReadByte
            Map(mapnum).Tile(x, y).Cubo = buffer.ReadByte
            Map(mapnum).Tile(x, y).HP = buffer.ReadLong
            Map(mapnum).Tile(x, y).Mensaje = buffer.ReadString
            Map(mapnum).Tile(x, y).Animacion = buffer.ReadLong
           'Para futuras versiones
           Map(mapnum).Tile(x, y).Banco = buffer.ReadLong
           Map(mapnum).Tile(x, y).Evento = buffer.ReadLong
           Map(mapnum).Tile(x, y).BancoLlave = buffer.ReadLong
           Map(mapnum).Tile(x, y).Script = buffer.ReadLong
           Map(mapnum).Tile(x, y).Timer = buffer.ReadLong
           Map(mapnum).Tile(x, y).ParteCubo = buffer.ReadByte
           Map(mapnum).Tile(x, y).SFX1 = buffer.ReadLong
           Map(mapnum).Tile(x, y).SFX2 = buffer.ReadLong
           Map(mapnum).Tile(x, y).Objeto = buffer.ReadLong

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).NPC(x) = buffer.ReadLong
        Map(mapnum).NpcSpawnType(x) = buffer.ReadLong
        Call ClearMapNpc(x, mapnum)
    Next
    
    'Event Data!
    Map(mapnum).EventCount = buffer.ReadLong
        
    If Map(mapnum).EventCount > 0 Then
        ReDim Map(mapnum).Events(0 To Map(mapnum).EventCount)
        For I = 1 To Map(mapnum).EventCount
            With Map(mapnum).Events(I)
                .Name = buffer.ReadString
                .Global = buffer.ReadLong
                .x = buffer.ReadLong
                .y = buffer.ReadLong
                .PageCount = buffer.ReadLong
            End With
            If Map(mapnum).Events(I).PageCount > 0 Then
                ReDim Map(mapnum).Events(I).Pages(0 To Map(mapnum).Events(I).PageCount)
                For x = 1 To Map(mapnum).Events(I).PageCount
                    With Map(mapnum).Events(I).Pages(x)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                            
                        .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                            
                        .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                        .HasItemAmount = buffer.ReadLong
                            
                        .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                            
                        .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                            
                        .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                            
                        .MoveRouteCount = buffer.ReadLong
                        
                        .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(mapnum).Events(I).Pages(x).MoveRoute(0 To .MoveRouteCount)
                            For y = 1 To .MoveRouteCount
                                .MoveRoute(y).index = buffer.ReadLong
                                .MoveRoute(y).Data1 = buffer.ReadLong
                                .MoveRoute(y).Data2 = buffer.ReadLong
                                .MoveRoute(y).Data3 = buffer.ReadLong
                                .MoveRoute(y).Data4 = buffer.ReadLong
                                .MoveRoute(y).Data5 = buffer.ReadLong
                                .MoveRoute(y).Data6 = buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .WalkThrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                            
                        .Position = buffer.ReadLong
                    End With
                        
                    If Map(mapnum).Events(I).Pages(x).CommandListCount > 0 Then
                        ReDim Map(mapnum).Events(I).Pages(x).CommandList(0 To Map(mapnum).Events(I).Pages(x).CommandListCount)
                        For y = 1 To Map(mapnum).Events(I).Pages(x).CommandListCount
                            Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount = buffer.ReadLong
                            Map(mapnum).Events(I).Pages(x).CommandList(y).ParentList = buffer.ReadLong
                            If Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount > 0 Then
                                ReDim Map(mapnum).Events(I).Pages(x).CommandList(y).Commands(1 To Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount)
                                For z = 1 To Map(mapnum).Events(I).Pages(x).CommandList(y).CommandCount
                                    With Map(mapnum).Events(I).Pages(x).CommandList(y).Commands(z)
                                        .index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .Data1 = buffer.ReadLong
                                        .Data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .Data5 = buffer.ReadLong
                                        .Data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.Data1 = buffer.ReadLong
                                        .ConditionalBranch.Data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).index = buffer.ReadLong
                                                .MoveRoute(w).Data1 = buffer.ReadLong
                                                .MoveRoute(w).Data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).Data5 = buffer.ReadLong
                                                .MoveRoute(w).Data6 = buffer.ReadLong
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    
    'End Event Data

    Call SendMapNpcsToMap(mapnum)
    Call SpawnMapNpcs(mapnum)
    Call SpawnGlobalEvents(mapnum)
    
    For I = 1 To Player_HighIndex
        If Player(I).Map = mapnum Then
            SpawnMapEventsFor I, mapnum
        End If
    Next

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).x, MapItem(GetPlayerMap(index), I).y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))
    ' Save the map
    Call SaveMap(mapnum)
    Call MapCache_Create(mapnum)
    Call ClearTempTile(mapnum)
    Call CacheResources(mapnum)

    ' Refresh map for everyone online
    For I = 1 To Player_HighIndex
        If isPlaying(I) And GetPlayerMap(I) = mapnum Then
            Call PlayerWarp(I, mapnum, GetPlayerX(I), GetPlayerY(I))
        End If
    Next I
    
    Call CacheMapBlocks(mapnum)

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim buffer As clsBuffer
    Dim I As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Get yes/no value
    s = buffer.ReadLong 'Parse(1)
    Set buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SpawnMapEventsFor(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    'send Resource cache
    For I = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, I
    Next

    TempPlayer(index).GettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call PlayerMapGetItem(index)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim invNum As Long
    Dim Amount As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    invNum = buffer.ReadLong 'CLng(Parse(1))
    Amount = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, invNum) < 1 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > GetPlayerInvItemValue(index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    If Item(GetPlayerInvItemNum(index, invNum)).BindType = 0 Then
    Call PlayerMapDropItem(index, invNum, Amount)
    Else
    PlayerMsg index, "No puedes tirar este objeto.", BrightRed
  End If
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), I).x, MapItem(GetPlayerMap(index), I).y)
        Call ClearMapItem(I, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNpc(I, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    If mensajeactualizar <> False Then
    Call PlayerMsg(index, "Mapa Actualizado.", Blue)
    End If
    
    mensajeactualizar = True
    Call AddLog(GetPlayerName(index) & " ha actualizado el mapa #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim I As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    s = "Mapas libres: "
    tMapStart = 1
    tMapEnd = 1

    For I = 1 To MAX_MAPS

        If LenB(Trim$(Map(I).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else

            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = I + 1
            tMapEnd = I + 1
        End If

    Next

    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            If GetPlayerAccess(N) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(N) & " ha sido expulsado de " & Options.Game_Name & " por " & GetPlayerName(index) & ".", White)
                Call AddLog(GetPlayerName(index) & " ha expulsado a " & GetPlayerName(N) & ".", ADMIN_LOG)
                Call AlertMsg(N, "Has sido expulsado por " & GetPlayerName(index) & ".")
            Else
                Call PlayerMsg(index, "El personaje tiene igual o mayor privilegio que tú.", White) ' MateoD
            End If

        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "No puedes expulsarte a ti mismo.", White) ' MateoD
    End If

End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    N = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(index, N & ": banneo de IP " & s & " por " & Name, White) ' MateoD
        N = N + 1
    Loop

    Close #F
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim filename As String
    Dim File As Long
    Dim F As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(index, "Todos los banneos han sido eliminados.", White) ' MateoD
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    Set buffer = Nothing

    If N <> index Then
        If N > 0 Then
            If GetPlayerAccess(N) < GetPlayerAccess(index) Then
                Call BanIndex(N, index)
            Else
                Call PlayerMsg(index, "El jugador tiene igual o mayor rango que tú.", White) ' MateoD
            End If

        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "No puedes bannearte a ti mismo.", White) ' MateoD
    End If

End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If
    
    SendMapEventData (index)

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SItemEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ItemSize = LenB(Item(N))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(N)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateItemToAll(N)
    Call SaveItem(N)
    Call AddLog(GetPlayerName(index) & " guardo el Objeto #" & N & ".", ADMIN_LOG)
End Sub

Sub HandleSaveCombo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim ComboSize As Long
    Dim ComboData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
    ComboSize = LenB(Combo(N))
    ReDim ComboData(ComboSize - 1)
    ComboData = buffer.ReadBytes(ComboSize)
    CopyMemory ByVal VarPtr(Combo(N)), ByVal VarPtr(ComboData(0)), ComboSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateComboToAll(N)
    Call SaveCombo(N)
    Call AddLog(GetPlayerName(index) & " guardo el Combo #" & N & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimationEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    N = buffer.ReadLong 'CLng(Parse(1))

    If N < 0 Or N > MAX_ANIMATIONS Then
        Exit Sub
    End If

    ' Update the Animation
    AnimationSize = LenB(Animation(N))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(N)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(N)
    Call SaveAnimation(N)
    Call AddLog(GetPlayerName(index) & " guardo la Animacion #" & N & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim npcNum As Long
    Dim buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    npcNum = buffer.ReadLong

    ' Prevent hacking
    If npcNum < 0 Or npcNum > MAX_NPCS Then
        Exit Sub
    End If

    NPCSize = LenB(NPC(npcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(npcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    ' Save it
    Call SendUpdateNpcToAll(npcNum)
    Call SaveNpc(npcNum)
    Call AddLog(GetPlayerName(index) & " guardo el NPC #" & npcNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " guardo el Recurso #" & ResourceNum & ".", ADMIN_LOG)
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SShopEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim shopNum As Long
    Dim I As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    shopNum = buffer.ReadLong

    ' Prevent hacking
    If shopNum < 0 Or shopNum > MAX_SHOPS Then
        Exit Sub
    End If

    ShopSize = LenB(Shop(shopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(shopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    ' Save it
    Call SendUpdateShopToAll(shopNum)
    Call SaveShop(shopNum)
    Call AddLog(GetPlayerName(index) & " guardo la Tienda #" & shopNum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditspell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SSpellEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim spellnum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    spellnum = buffer.ReadLong

    ' Prevent hacking
    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Exit Sub
    End If

    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(spellnum)), ByVal VarPtr(SpellData(0)), SpellSize
    ' Save it
    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)
    Call AddLog(GetPlayerName(index) & " guardo el Hechizo #" & spellnum & ".", ADMIN_LOG)
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The access
    I = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check for invalid access level
    If I >= 0 Or I <= 3 Then

        ' Check if player is on
        If N > 0 Then

            'check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(N) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Privilegios insuficientes.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(N) <= 0 Then
                Call GlobalMsg(GetPlayerName(N) & " ahora posee privilegios administrativos.", BrightBlue) ' MateoD
            End If

            Call SetPlayerAccess(N, I)
            Call SendPlayerData(N)
            Call AddLog(GetPlayerName(index) & " modifico a " & GetPlayerName(N) & "los Privilegios.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
        End If

    Else
        Call PlayerMsg(index, "Privilegios insuficientes.", Red)
    End If

End Sub

' :::::::::::::::::::::::
' :: Set name packet ::
' :::::::::::::::::::::::
Sub HandleSetName(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim I As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    ' The index
    N = FindPlayer(buffer.ReadString) 'Parse(1))
    ' The name
    I = buffer.ReadString 'CLng(Parse(2))
    Set buffer = Nothing

    ' Check if player is on
    If N > 0 Then

        'check to see if same level access is trying to change another access of the very same level and boot them if they are.
        If GetPlayerAccess(N) = GetPlayerAccess(index) Then
            Call PlayerMsg(index, "Privilegios insuficientes.", Red)
            Exit Sub
        End If
            
        Call AddLog(GetPlayerName(index) & " modifico a " & GetPlayerName(N) & " el nombre " & I & ".", ADMIN_LOG)
        Call SetPlayerName(N, I)
        Call SendPlayerData(N)
            
        If GetPlayerAccess(N) <= 0 Then
            Call PlayerMsg(N, "Tu nombre ha sido modificado.", White)
        End If
    Else
        Call PlayerMsg(index, "El jugador está desconectado", White) ' MateoD
    End If

End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(buffer.ReadString) 'Parse(1))
    SaveOptions
    Set buffer = Nothing
    Call GlobalMsg("Noticia cambiada a: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " cambio la Noticia a: " & Options.MOTD, ADMIN_LOG)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
On Error GoTo error:
    Dim x As Long
    Dim y As Long
    Dim I As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong 'CLng(Parse(1))
    y = buffer.ReadLong 'CLng(Parse(2))
    Set buffer = Nothing

    ' Prevent subscript out of range
    'If x < 0 Or x > Map(GetPlayerMap(index)).MaxX Or y < 0 Or y > Map(GetPlayerMap(index)).MaxY Then
    '   Exit Sub
   ' End If

    ' Check for a player
    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If GetPlayerMap(index) = GetPlayerMap(I) Then
                If Not GetPlayerVisible(I) = 1 Then
                    If GetPlayerX(I) = x Then
                        If GetPlayerY(I) = y Then
                        ' Change target
                            If TempPlayer(index).targetType = TARGET_TYPE_PLAYER And TempPlayer(index).target = I Then
                                TempPlayer(index).target = 0
                                TempPlayer(index).targetType = TARGET_TYPE_NONE
                                ' send target to player
                                SendTarget index
                            Else
                                TempPlayer(index).target = I
                                TempPlayer(index).targetType = TARGET_TYPE_PLAYER
                                ' send target to player
                                SendTarget index
                            End If
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For I = 1 To MAX_MAP_NPCS
        If MapNpc(GetPlayerMap(index)).NPC(I).num > 0 Then
            If MapNpc(GetPlayerMap(index)).NPC(I).x = x Then
                If MapNpc(GetPlayerMap(index)).NPC(I).y = y Then
                    If TempPlayer(index).target = I And TempPlayer(index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).target = 0
                        TempPlayer(index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget index
                    Else
                        ' Change target
                        TempPlayer(index).target = I
                        TempPlayer(index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for Spawn Tile
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_PLAYERSPAWN Then
        If GetPlayerX(index) = x Or GetPlayerX(index) + 1 = x Or GetPlayerX(index) - 1 = x Then ' Player is to west or east or on same X of spawn tile
            If GetPlayerY(index) = y Or GetPlayerY(index) + 1 = y Or GetPlayerY(index) - 1 = y Then ' Player is to south of north or on same Y of spawn tile
                SetPlayerSpawn index, GetPlayerMap(index), x, y
                PlayerMsg index, "Tu punto de spawn ha sido reiniciado.", Yellow ' MateoD
            End If
        End If
    End If
error:
End Sub

' :::::::::::::::::::
' : Location Packet :
' :::::::::::::::::::
Sub HandleBeFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim pI As Long
    Dim I As Long, II As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    pI = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    
        'make sure the friend's system is activated
        If Not frmServer.chkFriendSystem.Value = vbChecked Then Exit Sub

        If isPlaying(pI) Then
            
            ' If already friends, exit out
            For II = 1 To Player(index).Friends.Count
                If GetPlayerFriendName(index, II) = GetPlayerName(pI) Then Exit Sub
                'If GetPlayerFriendName(pI, II) = GetPlayerName(index) Then Exit Sub
            Next II
                            
            ' If player has max amount of friends, exit out
            If GetPlayerFriends(index) + 1 > MAX_FRIENDS Then
                Call PlayerMsg(index, "Lista de amigos llena.", BrightRed) ' MateoD
                Exit Sub
            End If
                            
            ' If clicked player has max amount of friends, exit out
            If GetPlayerFriends(pI) + 1 > MAX_FRIENDS Then
                Call PlayerMsg(index, GetPlayerName(pI) & " tiene la lista de amigos llena.", BrightRed) ' MateoD
                Exit Sub
            End If
            
            ' Make sure player hasn't reached friend request limit.
            If GetPlayerFriendRequests(index) + 1 > MAX_REQUESTS Then
                Call PlayerMsg(index, "Has enviado muchas peticiones sin respuesta. Espera 5 minutos.", BrightRed) ' MateoD
                Exit Sub
            End If
                            
            ' We're good, ask other player for friendship permission.
            Call SetPlayerFriendRequests(index, 1)
            Call AskForFriendshipFrom(pI, GetPlayerName(index))
            Call PlayerMsg(index, "Peticion de amistad enviada.", Orange) ' or maybe yellow ' MateoD
                            
        End If
End Sub

Sub HandleAcceptFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tempStr As String
Dim pI As Long, I As Long
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    tempStr = buffer.ReadString
    
    If Not Len(tempStr) > 0 Then
        Call PlayerMsg(index, "No puedes responder a la petición de amistad. Por favor vuelve a intentar.", BrightRed) ' MateoD
        Exit Sub
    End If
    
    pI = FindPlayer(tempStr)
    
    ' Go ahead and just tell the player we're good to go.
    Call PlayerMsg(pI, GetPlayerName(index) & " ha aceptado tu solicitud de amistad.", BrightGreen) ' MateoD

    ' We have permission, let's make these two buds.
    Call SetPlayerFriends(pI, 1)
            
    ' Update and tell the other player
    Call SetPlayerFriendName(pI, GetPlayerFriends(pI), GetPlayerName(index))
    Call PlayerMsg(pI, "Se ha incorporado a tu lista de amigos a " & GetPlayerName(index), Cyan) ' MateoD
    
    ' Subtract a request point
    If GetPlayerFriendRequests(pI) > 0 Then Call SetPlayerFriendRequests(pI, -1)
    
    'make sure we're not doubling up friends
    For I = 1 To GetPlayerFriends(index)
        If GetPlayerFriendName(index, I) = GetPlayerName(pI) Then GoTo SkipThatShit
    Next I
    
    ' Update and tell yourself
    Call SetPlayerFriendName(index, GetPlayerFriends(index), GetPlayerName(pI))
    Call PlayerMsg(index, "Se ha incorporado a tu lista de amigos a " & GetPlayerName(pI), Cyan) ' MateoD
    Call SetPlayerFriends(index, 1)
                            
SkipThatShit:
    ' Send new data to both players
    Call SendDataTo(index, PlayerFriends(index))
    Call SendDataTo(pI, PlayerFriends(pI))
    
    Set buffer = Nothing
End Sub

Sub HandleRequestFriendData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tempStr As String, pI As Long, I As Long
Dim pData(1 To 6) As String
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tempStr = buffer.ReadString
    Set buffer = Nothing
    
    'Make sure we have a name
    If Not Len(tempStr) > 0 Then Exit Sub
    pI = FindPlayer(tempStr)
    
    'Make sure we have an index
    If pI < 1 Or pI > MAX_PLAYERS Then Exit Sub
    
    
    'Start setting data
    pData(1) = GetPlayerLevel(pI)
    pData(2) = GetPlayerStat(pI, strength)
    pData(3) = GetPlayerStat(pI, Endurance)
    pData(4) = GetPlayerStat(pI, Intelligence)
    pData(5) = GetPlayerStat(pI, Agility)
    pData(6) = GetPlayerStat(pI, Willpower)
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFriendData
    For I = 1 To UBound(pData)
        buffer.WriteLong pData(I)
    Next I
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleDeclineFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tempStr As String
Dim pI As Long
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    tempStr = buffer.ReadString
    
    If Not Len(tempStr) > 0 Then
        Call PlayerMsg(index, "No puedes responder a la solicitud de amistad. Por favor vuelve a intentar.", BrightRed) ' MateoD
        Exit Sub
    End If
    
    pI = FindPlayer(tempStr)
    ' Simply tell the player the request was declined.
    Call PlayerMsg(pI, GetPlayerName(index) & " ha rechazado tu solicitud de amistad.", BrightRed) ' MateoD
    
    ' Subtract a request point (On second thought, no. lol)
    'If GetPlayerFriendRequests(index) > 0 Then Call SetPlayerFriendRequests(index, -1)
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Spell slot
    N = buffer.ReadLong 'CLng(Parse(1))
    Set buffer = Nothing
    ' set the spell buffer before castin
    Call BufferSpell(index, N)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call CloseSocket(index)
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, N As Long
    
    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).spellBuffer.Spell > 0 Then
        PlayerMsg index, "Imposible, estás lanzando el hechizo.", BrightRed ' MateoD
        Exit Sub
    End If
    
    For N = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(N) > GetTickCount Then
            PlayerMsg index, "Imposible, el hechizo está en enfriamiento.", BrightRed ' MateoD
            Exit Sub
        End If
    Next
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    Set buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PlayerUnequipItem index, buffer.ReadLong
    Set buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestCombos(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendCombos index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleRequestEditCombos(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong SComboEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' item
    tmpItem = buffer.ReadLong
    tmpAmount = buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim thePlr As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    thePlr = FindPlayer(buffer.ReadString)

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    SetPlayerExp thePlr, GetPlayerNextLevel(thePlr)
    CheckPlayerLevelUp thePlr
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim spellslot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    spellslot = buffer.ReadLong
    
    ' Check for subscript out of range
    If spellslot < 1 Or spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a spell which is in CD
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        PlayerMsg index, "Imposible, el hechizo está en enfriamiento.", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a spell which is buffered
    If TempPlayer(index).spellBuffer.Spell = spellslot Then
        PlayerMsg index, "Imposible, el hechizo está en lanzamiento.", BrightRed
        Exit Sub
    End If
    
    Player(index).Spell(spellslot) = 0
    SendPlayerSpells index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim shopslot As Long
    Dim shopNum As Long
    Dim itemAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    shopslot = buffer.ReadLong
    
    ' not in shop, exit out
    shopNum = TempPlayer(index).InShop
    If shopNum < 1 Or shopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(shopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemAmount = HasItem(index, .costitem)
        If itemAmount = 0 Or itemAmount < .costvalue Then
            PlayerMsg index, "Dinero insuficiente.", BrightRed
            ResetShopAction index
            Exit Sub
        End If
        
        ' it's fine, let's go ahead
        TakeInvItem index, .costitem, .costvalue
        GiveInvItem index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Transacción realizada.", BrightGreen ' MateoD
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim itemnum As Long
    Dim price As Long
    Dim multiplier As Double
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    itemnum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    price = Item(itemnum).price * multiplier
    
    ' item has cost?
    If price <= 0 Then
        PlayerMsg index, "El vendedor no quiere este objeto.", BrightRed ' MateoD
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, itemnum, 1
    GiveInvItem index, 1, price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Transacción realizada.", BrightGreen ' MateoD
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    oldSlot = buffer.ReadLong
    newSlot = buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set buffer = Nothing
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Long
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    BankSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    TakeBankItem index, BankSlot, Amount
    
    Set buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    GiveBankItem index, invSlot, Amount
    
    Set buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    x = buffer.ReadLong
    y = buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        'PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    
    Set buffer = Nothing
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    ' can't trade npcs
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).target
    
    ' make sure we don't error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can't trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "No puedes comerciar contigo mismo.", BrightRed
        Exit Sub
    End If
    
    ' make sure they're on the same map
    If Not Player(tradeTarget).Map = Player(index).Map Then Exit Sub
    
    ' make sure they're stood next to each other
    tX = Player(tradeTarget).x
    tY = Player(tradeTarget).y
    sX = Player(index).x
    sY = Player(index).y
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "Debes estar próximo al objetivo para enviar una petición de comercio.", BrightRed  ' MateoD
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "Debes estar próximo al objetivo para enviar una petición de comercio.", BrightRed  ' MateoD
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "El jugador está ocupado.", BrightRed  ' MateoD
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendTradeRequest tradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim tradeTarget As Long
Dim I As Long

    If TempPlayer(index).InTrade > 0 Then
        TempPlayer(index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(index).TradeRequest
        ' let them know they're trading
        PlayerMsg index, "La solicitud de " & Trim$(GetPlayerName(tradeTarget)) & " ha sido aceptada.", BrightGreen
        PlayerMsg tradeTarget, Trim$(GetPlayerName(index)) & " ha aceptado tu solicitud.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they're trading with each other
        TempPlayer(index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = index
        ' clear out their trade offers
        For I = 1 To MAX_INV
            TempPlayer(index).TradeOffer(I).num = 0
            TempPlayer(index).TradeOffer(I).Value = 0
            TempPlayer(tradeTarget).TradeOffer(I).num = 0
            TempPlayer(tradeTarget).TradeOffer(I).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade index, tradeTarget
        SendTrade tradeTarget, index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate index, 0
        SendTradeUpdate index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " ha rechazado tu solicitud.", BrightRed ' MateoD
    PlayerMsg index, "Has rechazado la solicitud.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim tradeTarget As Long
    Dim I As Long
    Dim tmpTradeItem(1 To MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(1 To MAX_INV) As PlayerInvRec
    Dim itemnum As Long
    
    TempPlayer(index).AcceptTrade = True
    
    tradeTarget = TempPlayer(index).InTrade
        
    If tradeTarget > 0 Then
    
        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' take their items
        For I = 1 To MAX_INV
            ' player
            If TempPlayer(index).TradeOffer(I).num > 0 Then
                itemnum = Player(index).Inv(TempPlayer(index).TradeOffer(I).num).num
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem(I).num = itemnum
                    tmpTradeItem(I).Value = TempPlayer(index).TradeOffer(I).Value
                    ' take item
                    TakeInvSlot index, TempPlayer(index).TradeOffer(I).num, tmpTradeItem(I).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(I).num > 0 Then
                itemnum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num)
                If itemnum > 0 Then
                    ' store temp
                    tmpTradeItem2(I).num = itemnum
                    tmpTradeItem2(I).Value = TempPlayer(tradeTarget).TradeOffer(I).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(I).num, tmpTradeItem2(I).Value
                End If
            End If
        Next
    
        ' taken all items. now they can't not get items because of no inventory space.
        For I = 1 To MAX_INV
            ' player
            If tmpTradeItem2(I).num > 0 Then
                ' give away!
                GiveInvItem index, tmpTradeItem2(I).num, tmpTradeItem2(I).Value, False
            End If
            ' target
            If tmpTradeItem(I).num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(I).num, tmpTradeItem(I).Value, False
            End If
        Next
    
        SendInventory index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For I = 1 To MAX_INV
            TempPlayer(index).TradeOffer(I).num = 0
            TempPlayer(index).TradeOffer(I).Value = 0
            TempPlayer(tradeTarget).TradeOffer(I).num = 0
            TempPlayer(tradeTarget).TradeOffer(I).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "Transacción completa.", BrightGreen ' MateoD
        PlayerMsg tradeTarget, "Transacción completa.", BrightGreen ' MateoD
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
            
    End If
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim I As Long
Dim tradeTarget As Long

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then
        For I = 1 To MAX_INV
            TempPlayer(index).TradeOffer(I).num = 0
            TempPlayer(index).TradeOffer(I).Value = 0
            TempPlayer(tradeTarget).TradeOffer(I).num = 0
            TempPlayer(tradeTarget).TradeOffer(I).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
    
        PlayerMsg index, "Has rechazado la oferta.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(index) & " ha rechazado tu oferta.", BrightRed ' MateoD
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
    End If
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim invSlot As Long
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim itemnum As Long
    Dim I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    invSlot = buffer.ReadLong
    Amount = buffer.ReadLong
    
    Set buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(index, invSlot)
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If Amount < 0 Or Amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If

    If Item(itemnum).Type = ITEM_TYPE_CURRENCY Or Item(itemnum).Stackable > 0 Then
        ' check if already offering same currency item
        For I = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(I).num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(I).Value = TempPlayer(index).TradeOffer(I).Value + Amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(I).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(I).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they're not already offering it
        For I = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(I).num = invSlot Then
                PlayerMsg index, "Ya has ofrecido este objeto.", BrightRed ' MateoD
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For I = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(I).num = 0 Then
            EmptySlot = I
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = Amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim tradeSlot As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    tradeSlot = buffer.ReadLong
    
    Set buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim sType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    sType = buffer.ReadLong
    Slot = buffer.ReadLong
    hotbarNum = buffer.ReadLong
    
    Select Case sType
        Case 0 ' clear
            Player(index).Hotbar(hotbarNum).Slot = 0
            Player(index).Hotbar(hotbarNum).sType = 0
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If Player(index).Inv(Slot).num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Inv(Slot).num
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
        Case 2 ' spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Player(index).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Player(index).Spell(Slot)).Name)) > 0 Then
                        Player(index).Hotbar(hotbarNum).Slot = Player(index).Spell(Slot)
                        Player(index).Hotbar(hotbarNum).sType = sType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set buffer = Nothing
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long
    Dim I As Long
    Dim tipo As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadLong
    
    'easse canido
    Select Case Player(index).Hotbar(Slot).sType
        Case 1 ' inventory
            For I = 1 To MAX_INV
                If Player(index).Inv(I).num > 0 Then
                    If Player(index).Inv(I).num = Player(index).Hotbar(Slot).Slot Then
                        If Item(Player(index).Inv(I).num).Type = ITEM_TYPE_CONSUME Then
                            'Elimina los items consumibles que no son apilables
                            If Not Item(Player(index).Inv(I).num).Stackable = 1 Then
                                Player(index).Hotbar(Slot).Slot = 0
                                Player(index).Hotbar(Slot).sType = 0
                            End If
                            'elimina ultimo items consumible apilable
                            If Player(index).Inv(I).Value = 1 Then
                                Player(index).Hotbar(Slot).Slot = 0
                                Player(index).Hotbar(Slot).sType = 0
                            End If
                            SendHotbar index
                        End If
                        UseItem index, I
                        Exit Sub
                    End If
                End If
                
                'si el objeto esta equipado entonces lo desequipa(asi se escribe? soy iki)
                For tipo = 1 To 4
                       If GetPlayerEquipment(index, tipo) = Player(index).Hotbar(Slot).Slot Then
                            PlayerUnequipItem index, tipo
                            SendHotbar index
                            Exit Sub
                       End If
                Next
                            
                 
            Next
        Case 2 ' spell
            For I = 1 To MAX_PLAYER_SPELLS
                If Player(index).Spell(I) > 0 Then
                    If Player(index).Spell(I) = Player(index).Hotbar(Slot).Slot Then
                        BufferSpell index, I
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set buffer = Nothing
End Sub
Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' make sure it's a valid target
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).target = index Then Exit Sub
    
    ' make sure they're connected and on the same map
    If Not IsConnected(TempPlayer(index).target) Or Not isPlaying(TempPlayer(index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).target
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(index).partyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).partyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandleEventChatReply(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim eventID As Long, pageID As Long, reply As Long, I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    eventID = buffer.ReadLong
    pageID = buffer.ReadLong
    reply = buffer.ReadLong
    
    If TempPlayer(index).EventProcessingCount > 0 Then
        For I = 1 To TempPlayer(index).EventProcessingCount
            If TempPlayer(index).EventProcessing(I).eventID = eventID And TempPlayer(index).EventProcessing(I).pageID = pageID Then
                If TempPlayer(index).EventProcessing(I).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).index = EventType.evShowText Then
                            TempPlayer(index).EventProcessing(I).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(index).EventProcessing(I).ListLeftOff(TempPlayer(index).EventProcessing(I).CurList) = TempPlayer(index).EventProcessing(I).CurSlot
                                    TempPlayer(index).EventProcessing(I).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).Data1
                                    TempPlayer(index).EventProcessing(I).CurSlot = 1
                                Case 2
                                    TempPlayer(index).EventProcessing(I).ListLeftOff(TempPlayer(index).EventProcessing(I).CurList) = TempPlayer(index).EventProcessing(I).CurSlot
                                    TempPlayer(index).EventProcessing(I).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).Data2
                                    TempPlayer(index).EventProcessing(I).CurSlot = 1
                                Case 3
                                    TempPlayer(index).EventProcessing(I).ListLeftOff(TempPlayer(index).EventProcessing(I).CurList) = TempPlayer(index).EventProcessing(I).CurSlot
                                    TempPlayer(index).EventProcessing(I).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).Data3
                                    TempPlayer(index).EventProcessing(I).CurSlot = 1
                                Case 4
                                    TempPlayer(index).EventProcessing(I).ListLeftOff(TempPlayer(index).EventProcessing(I).CurList) = TempPlayer(index).EventProcessing(I).CurSlot
                                    TempPlayer(index).EventProcessing(I).CurList = Map(GetPlayerMap(index)).Events(eventID).Pages(pageID).CommandList(TempPlayer(index).EventProcessing(I).CurList).Commands(TempPlayer(index).EventProcessing(I).CurSlot - 1).Data4
                                    TempPlayer(index).EventProcessing(I).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(index).EventProcessing(I).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    
    
    
    Set buffer = Nothing
End Sub

Sub HandleEvent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim N As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long, begineventprocessing As Boolean, z As Long, buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    I = buffer.ReadLong
    Set buffer = Nothing
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
            If TempPlayer(index).EventMap.EventPages(z).eventID = I Then
                I = z
                begineventprocessing = True
                Exit For
            End If
        Next
    End If
    
    If begineventprocessing = True Then
        If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).CommandListCount > 0 Then
            'Process this event, it is action button and everything checks out.
            TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
            With TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                .ActionTimer = GetTickCount
                .CurList = 1
                .CurSlot = 1
                .eventID = TempPlayer(index).EventMap.EventPages(I).eventID
                .pageID = TempPlayer(index).EventMap.EventPages(I).pageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(I).eventID).Pages(TempPlayer(index).EventMap.EventPages(I).pageID).CommandListCount)
            End With
            'Call CheckTasks(index, QUEST_TYPE_GOGETFROMEVENT, TempPlayer(index).EventMap.EventPages(i).eventID)
        End If
        begineventprocessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (index)
End Sub

Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, I As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For I = 1 To MAX_SWITCHES
        Switches(I) = buffer.ReadString
    Next
    
    For I = 1 To MAX_VARIABLES
        Variables(I) = buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

Sub HandlePlayerVisibility(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

    If Not Player(index).Visible = 0 Then
        Player(index).Visible = 0
    Else
        Player(index).Visible = 1
    End If
    
    Call SendPlayerData(index)
End Sub

' ::::::::::::::::::::::::
' :: Heal Player packet ::
' ::::::::::::::::::::::::
Sub HandleHealPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    N = FindPlayer(buffer.ReadString)
    Set buffer = Nothing

    ' Check if player is on
    If N > 0 Then
        Call SetPlayerVital(N, Vitals.HP, GetPlayerMaxVital(N, Vitals.HP))
        Call SetPlayerVital(N, Vitals.MP, GetPlayerMaxVital(N, Vitals.MP))
        Call SendVital(N, Vitals.HP)
        Call SendVital(N, Vitals.MP)
        Call PlayerMsg(N, "Has sido curado por " & GetPlayerName(index) & ".", BrightBlue) ' MateoD
        Call AddLog(GetPlayerName(index) & " curo a" & GetPlayerName(N) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
    End If

End Sub

' ::::::::::::::::::::::::
' :: Kill Player packet ::
' ::::::::::::::::::::::::
Sub HandleKillPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim N As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MONITOR Then Exit Sub

    ' The index
    N = FindPlayer(buffer.ReadString)
    Set buffer = Nothing

    ' Check if player is on
    If N > 0 Then
        Call SetPlayerVital(N, Vitals.HP, 0)
        Call SendVital(N, Vitals.HP)
        Call OnDeath(N)
        Call PlayerMsg(N, "Has sido asesinado por " & GetPlayerName(index) & ".", BrightRed) ' MateoD
        Call AddLog(GetPlayerName(index) & " aniquilo a" & GetPlayerName(N) & ".", ADMIN_LOG)
    Else
        Call PlayerMsg(index, "El jugador está desconectado.", White) ' MateoD
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Client Character Editor ::
' :::::::::::::::::::::::::::::
Sub SendCharEditorRequest(ByVal I As Long, ByVal command As Byte, ByVal num As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SCharEditorRequest
    
    Select Case command
        Case 1:
            buffer.WriteByte command
            buffer.WriteLong GetPlayerLevel(I)
            buffer.WriteLong GetPlayerExp(I)
            buffer.WriteLong GetPlayerPOINTS(I)
            buffer.WriteLong GetPlayerStat(I, Endurance)
            buffer.WriteLong GetPlayerStat(I, strength)
            buffer.WriteLong GetPlayerStat(I, Intelligence)
            buffer.WriteLong GetPlayerStat(I, Agility)
            buffer.WriteLong GetPlayerStat(I, Willpower)
            buffer.WriteByte GetPlayerCombatLevel(I, num)
            buffer.WriteLong GetPlayerCombatExp(I, num)
            buffer.WriteLong GetPlayerInvItemNum(I, num)
            buffer.WriteLong GetPlayerInvItemValue(I, num)
            buffer.WriteLong GetPlayerBankItemNum(I, num)
            buffer.WriteLong GetPlayerBankItemValue(I, num)
            buffer.WriteLong GetPlayerLevel(I)
        Case 2:
            buffer.WriteByte command
            buffer.WriteByte GetPlayerCombatLevel(I, num)
            buffer.WriteLong GetPlayerCombatExp(I, num)
        Case 3:
            buffer.WriteByte command
            buffer.WriteLong GetPlayerInvItemNum(I, num)
            buffer.WriteLong GetPlayerInvItemValue(I, num)
        Case 4:
            buffer.WriteByte command
            buffer.WriteLong GetPlayerBankItemNum(I, num)
            buffer.WriteLong GetPlayerBankItemValue(I, num)
    End Select
    
    SendDataTo I, buffer.ToArray
    Set buffer = Nothing
End Sub

Sub HandleCharEditorCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long, N As Long, command As Byte, plExp As Long, plPts As Long, pStr As Long, pEnd As Long, pInt As Long, pAgi As Long, pWill As Long
    Dim lvl As Long, invNum As Long, itmNum As Long, itmQty As Long, bnkNum As Long, bankNum As Byte, bankQty As Long
    Dim comType As Byte, comLvl As Byte, comExp As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    command = buffer.ReadByte
    
    Select Case command
        Case 1
            I = FindPlayer(buffer.ReadString)
            If Not I = 0 Then
                SendCharEditorRequest I, 1, 1
            Else
                Call PlayerMsg(index, "El jugador está desconectado.", AlertColor) ' MateoD
            End If
        Case 2
            I = FindPlayer(buffer.ReadString)
            lvl = buffer.ReadLong
            plExp = buffer.ReadLong
            plPts = buffer.ReadLong
            If GetPlayerLevel(I) < lvl Then
                SetPlayerPOINTS I, GetPlayerPOINTS(I) + (3 * (lvl - GetPlayerLevel(I)))
                SetPlayerLevel I, lvl
            Else
                SetPlayerLevel I, lvl
                SetPlayerExp I, plExp
                SetPlayerPOINTS I, plPts
            End If
            pEnd = buffer.ReadLong
            pStr = buffer.ReadLong
            pInt = buffer.ReadLong
            pAgi = buffer.ReadLong
            pWill = buffer.ReadLong
            If pEnd > 100 Then pEnd = 100
            If pStr > 100 Then pStr = 100
            If pInt > 100 Then pInt = 100
            If pAgi > 100 Then pAgi = 100
            If pWill > 100 Then pWill = 100
            SetPlayerStat I, Endurance, pEnd
            SetPlayerStat I, strength, pStr
            SetPlayerStat I, Intelligence, pInt
            SetPlayerStat I, Agility, pAgi
            SetPlayerStat I, Willpower, pWill
                invNum = buffer.ReadLong
                itmNum = buffer.ReadLong
                itmQty = buffer.ReadLong
            SetPlayerInvItemNum I, invNum, itmNum
            SetPlayerInvItemValue I, invNum, itmQty
                bnkNum = buffer.ReadLong
                bankNum = buffer.ReadLong
                bankQty = buffer.ReadLong
            SetPlayerBankItemNum I, bnkNum, bankNum
            SetPlayerBankItemValue I, bnkNum, bankQty
            SendInventoryUpdate I, invNum
            SendEXP I
            CheckPlayerLevelUp I
            SaveBank I
            SavePlayer I
            SendPlayerData I
                SendCharEditorRequest I, 1, 1
        Case 3
            I = FindPlayer(buffer.ReadString)
            comType = buffer.ReadByte
            If comType > MAX_COMBAT Then
                Call PlayerMsg(index, "Valor demasiado alto, solo se permite hasta " & MAX_COMBAT, AlertColor)
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not comType = 0 Then
                    SendCharEditorRequest I, 2, comType
                Else
                    Call PlayerMsg(index, "La habilidad de combate debe tener un valor mayor a 0.", AlertColor) ' MateoD
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor) ' MateoD
            End If
        Case 4
            I = FindPlayer(buffer.ReadString)
            comType = buffer.ReadByte
            comLvl = buffer.ReadByte
            comExp = buffer.ReadLong
            If comType > MAX_COMBAT Then
                Call PlayerMsg(index, "Valor demasiado alto, solo se permite hasta " & MAX_COMBAT, AlertColor)
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not comType = 0 Then
                    SetPlayerCombatLevel I, comLvl, comType
                    SetPlayerCombatExp I, comType, comExp
                    SendPlayerData I
                    SavePlayer I
                    SendCharEditorRequest I, 2, comType
                Else
                    Call PlayerMsg(index, "La habilidad de combate debe tener un valor mayor a 0.", AlertColor) ' MateoD
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor) ' MateoD
            End If
        Case 5
            I = FindPlayer(buffer.ReadString)
            N = buffer.ReadLong
            If N > MAX_INV Then
                Call PlayerMsg(index, "Valor demasiado alto, el máximo permitido es " & MAX_INV, AlertColor)
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not N = 0 Then
                    SendCharEditorRequest I, 3, N
                Else
                    Call PlayerMsg(index, "El valor de objeto debe ser mayor a 0.", AlertColor)
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor) ' MateoD
            End If
        Case 6
            I = FindPlayer(buffer.ReadString)
            invNum = buffer.ReadLong
            itmNum = buffer.ReadLong
            itmQty = buffer.ReadLong
            If invNum > MAX_INV Then
                Call PlayerMsg(index, "Valor demasiado alto, el máximo permitido es " & MAX_INV, AlertColor) ' MateoD
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not invNum = 0 Then
                    SetPlayerInvItemNum I, invNum, itmNum
                    SetPlayerInvItemValue I, invNum, itmQty
                    SendInventoryUpdate I, invNum
                    SendPlayerData I
                    SavePlayer I
                    SendCharEditorRequest I, 3, invNum
                Else
                    Call PlayerMsg(index, "El valor del objeto debe ser mayor a 0.", AlertColor) ' MateoD
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor)
            End If
        Case 7
            I = FindPlayer(buffer.ReadString)
            N = buffer.ReadLong
            If N > MAX_BANK Then
                Call PlayerMsg(index, "Valor demasiado alto, el máximo permitido es " & MAX_BANK, AlertColor) ' MateoD
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not N = 0 Then
                    SendCharEditorRequest I, 4, N
                Else
                    Call PlayerMsg(index, "El valor del objeto debe ser mayor a 0.", AlertColor) ' MateoD
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor) ' MateoD
            End If
        Case 8
            I = FindPlayer(buffer.ReadString)
            bnkNum = buffer.ReadLong
            bankNum = buffer.ReadLong
            bankQty = buffer.ReadLong
            If N > MAX_BANK Then
                Call PlayerMsg(index, "Valor demasiado alto, el máximo permitido es " & MAX_BANK, AlertColor) ' MateoD
                Exit Sub
            End If
        
            If Not I = 0 Then
                If Not bnkNum = 0 Then
                    SetPlayerBankItemNum I, bnkNum, bankNum
                    SetPlayerBankItemValue I, bnkNum, bankQty
                    SaveBank I
                    SavePlayer I
                    SendPlayerData I
                    SendCharEditorRequest I, 4, bankNum
                Else
                    Call PlayerMsg(index, "El valor del objeto debe ser mayor a 0.", AlertColor) ' MateoD
                End If
            Else
                Call PlayerMsg(index, "Jugador no encontrado.", AlertColor) ' MateoD
            End If
    End Select
End Sub
Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong SQuestEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim N As Long
Dim buffer As clsBuffer
Dim QuestSize As Long
Dim QuestData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
    Exit Sub
    End If
    
    N = buffer.ReadLong 'CLng(Parse(1))
    
    If N < 0 Or N > MAX_QUESTS Then
    Exit Sub
    End If
    
    ' Update the Quest
    QuestSize = LenB(Quest(N))
    ReDim QuestData(QuestSize - 1)
    QuestData = buffer.ReadBytes(QuestSize)
    CopyMemory ByVal VarPtr(Quest(N)), ByVal VarPtr(QuestData(0)), QuestSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(N)
    Call SaveQuest(N)
    Call AddLog(GetPlayerName(index) & " guardo la Mision #" & N & ".", ADMIN_LOG)
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    
End Sub

Sub HandlePlayerHandleQuest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim QuestNum As Long, Order As Long, I As Long, N As Long
Dim RemoveStartItems As Boolean

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    QuestNum = buffer.ReadLong
    'prevent error, but tell me about it QUICKCHANGE
    If QuestNum < 1 Then
        Call PlayerMsg(index, "No puedes acceder a los datos de la misión.", Red) ' MateoD
        Exit Sub
    End If
    Order = buffer.ReadLong '1 = accept quest, 2 = cancel quest

    If Order = 1 Then
        RemoveStartItems = False
        'Alatar v1.2
        For I = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(I).Item > 0 Then
                If FindOpenInvSlot(index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                    PlayerMsg index, "Necesitas espacio en el inventario.", BrightRed ' MateoD
                    RemoveStartItems = True
                    Exit For
                Else
                    If Item(Quest(QuestNum).GiveItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                        GiveInvItem index, Quest(QuestNum).GiveItem(I).Item, Quest(QuestNum).GiveItem(I).Value
                    Else
                        For N = 1 To Quest(QuestNum).GiveItem(I).Value
                            If FindOpenInvSlot(index, Quest(QuestNum).GiveItem(I).Item) = 0 Then
                                PlayerMsg index, "Necesitas espacio en el inventario.", BrightRed ' MateoD
                                RemoveStartItems = True
                                Exit For
                            Else
                                GiveInvItem index, Quest(QuestNum).GiveItem(I).Item, 1
                            End If
                        Next
                    End If
                End If
            End If
        Next

        If RemoveStartItems = False Then 'this means everything went ok
            Player(index).PlayerQuest(QuestNum).Status = QUEST_STARTED '1
            Player(index).PlayerQuest(QuestNum).ActualTask = 1
            Player(index).PlayerQuest(QuestNum).CurrentCount = 0
            PlayerMsg index, "Misión aceptada: " & Trim$(Quest(QuestNum).Name) & ".", BrightGreen ' MateoD
        End If
    
    ElseIf Order = 2 Then
        Player(index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED '2
        Player(index).PlayerQuest(QuestNum).ActualTask = 1
        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
        RemoveStartItems = True 'avoid exploits
        'AlertMsg (index), "Mision abandonada."
        PlayerMsg (index), "Mision abandonada.", Green
        'PlayerMsg index, Trim$(Quest(QuestNum).Name) & " ha sido abortada.", BrightGreen ' MateoD
    End If

    If RemoveStartItems = True Then
        For I = 1 To MAX_QUESTS_ITEMS
            If Quest(QuestNum).GiveItem(I).Item > 0 Then
                If HasItem(index, Quest(QuestNum).GiveItem(I).Item) > 0 Then
                    If Item(Quest(QuestNum).GiveItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem index, Quest(QuestNum).GiveItem(I).Item, Quest(QuestNum).GiveItem(I).Value
                    Else
                        For N = 1 To Quest(QuestNum).GiveItem(I).Value
                            TakeInvItem index, Quest(QuestNum).GiveItem(I).Item, 1
                        Next
                        
                    End If
                End If
            End If
        Next
    End If


    SavePlayer index
    SendPlayerData index
    SendPlayerQuests index
    
    Set buffer = Nothing
End Sub

Sub HandleQuestLogUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerQuests index
End Sub

Private Sub HandleProjecTileAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim curProjecTile As Long, I As Long, CurEquipment As Long

    ' prevent subscript
    If index > MAX_PLAYERS Or index < 1 Then Exit Sub
    
    ' get the players current equipment
    CurEquipment = GetPlayerEquipment(index, Weapon)
    
    ' check if they've got equipment
    If CurEquipment < 1 Or CurEquipment > MAX_ITEMS Then Exit Sub
    
    ' set the curprojectile
    For I = 1 To MAX_PLAYER_PROJECTILES
        If TempPlayer(index).ProjecTile(I).Pic = 0 Then
            ' just incase there is left over data
            ClearProjectile index, I
            ' set the curprojtile
            curProjecTile = I
            Exit For
        End If
    Next
    
    ' check for subscript
    If curProjecTile < 1 Then Exit Sub
    
    ' populate the data in the player rec
    With TempPlayer(index).ProjecTile(curProjecTile)
        .Damage = Item(CurEquipment).ProjecTile.Damage
        .Direction = GetPlayerDir(index)
        .Pic = Item(CurEquipment).ProjecTile.Pic
        .Range = Item(CurEquipment).ProjecTile.Range
        .Speed = Item(CurEquipment).ProjecTile.Speed
        .x = GetPlayerX(index)
        .y = GetPlayerY(index)
        .Municion = Item(CurEquipment).ProjecTile.Municion
    End With
                
    ' trololol, they have no more projectile space left
    If curProjecTile < 1 Or curProjecTile > MAX_PLAYER_PROJECTILES Then Exit Sub
    
    ' update the projectile on the map
    SendProjectileToMap index, curProjecTile
    
End Sub

Sub ProcesarCuboMapa(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'FUNCION PARA CUBOS 32x32
    Dim tipo, Capa As Long
    Dim mapa As Long, Mapax As Long, Mapay As Long
    Dim TileX As Long, TileY As Long, TileN As Long
    Dim buffer As clsBuffer
    Dim I As Integer
    Dim direccion As Long
    Dim Equipamiento, Inventario, Cantidad As Long
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    mensajeactualizar = False 'Para que no nos salga el mensaje de mapa actualizado
    
'Recibimos datos y asignamos a variables

TileX = buffer.ReadLong
TileY = buffer.ReadLong
TileN = buffer.ReadLong
Capa = buffer.ReadLong
Mapax = buffer.ReadLong
Mapay = buffer.ReadLong
mapa = buffer.ReadLong
tipo = buffer.ReadLong


Select Case GetPlayerDir(index) 'Anulamos el proceso si hay limites

Case DIR_UP

If GetPlayerY(index) = 0 Then Exit Sub

Case DIR_DOWN

If GetPlayerY(index) = Map(mapa).MaxY Then Exit Sub

Case DIR_LEFT

If GetPlayerX(index) = 0 Then Exit Sub

Case DIR_RIGHT

If GetPlayerX(index) = Map(mapa).MaxX Then Exit Sub

End Select


'CUBO 1


'Tile CUBO1
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = True



    Call SpawnMapItems(mapa)


    ' Guardar y limpiar
    Call SaveMap(mapa)
    Call MapCache_Create(mapa)
    Call ClearTempTile(mapa)
    
    ' Enviar a todos
    For I = 1 To Player_HighIndex
        If isPlaying(I) And GetPlayerMap(I) = mapa Then
            Call PlayerWarp(I, mapa, GetPlayerX(I), GetPlayerY(I))
            Call SendMapItemsToAll(I)
        End If
    Next I
    
     Call CacheMapBlocks(mapa)
    
    Set buffer = Nothing
    
   
'Inventario EaSee Engine 0.4 Version Cubos by ikicapitan (Podria utilizarse para sistema de flechas inclusive)
Equipamiento = GetPlayerEquipment(index, Weapon)
Inventario = 0

For I = 1 To MAX_INV 'Busca en el inventario
Inventario = GetPlayerInvItemNum(index, I)
Cantidad = GetPlayerInvItemValue(index, I)
If Inventario = Equipamiento Then 'Si hay un objeto que coincida con el objeto equipado
TakeInvItem index, Inventario, 1  ' Toma el objeto del inventario en cantidad 1

GoTo ActualizarEquipamiento
End If

Next

SetPlayerEquipment index, 0, Weapon 'Sino lo quita del equipado

ActualizarEquipamiento:

SendMapEquipment index
SendInventory index

   
End Sub


Sub ProcesarCuboMapa64(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'FUNCION PARA CUBOS 32x64
    'CUBO1
    Dim tipo, Capa As Long
    Dim mapa As Long, Mapax As Long, Mapay As Long
    Dim TileX As Long, TileY As Long, TileN As Long
    'CUBO2
    Dim Tipo2, Capa2 As Long
    Dim Mapax2, Mapay2 As Long
    Dim TileX2, TileY2 As Long
    'Recicladores
    Dim A As Long, b As String, C As Byte
    'Inventario y Equipamiento
    Dim Inventario, Equipamiento, Cantidad As Long
    
    Dim buffer As clsBuffer
    Dim I As Integer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    mensajeactualizar = False 'Para que no nos salga el mensaje de mapa actualizado
'Recibimos datos y asignamos a variables


TileX = buffer.ReadLong
TileY = buffer.ReadLong
TileN = buffer.ReadLong
Capa = buffer.ReadLong
Mapax = buffer.ReadLong
Mapay = buffer.ReadLong
mapa = buffer.ReadLong
tipo = buffer.ReadLong





'Tile CUBO2

Select Case GetPlayerDir(index)

Case DIR_UP

'Cubo 1
If GetPlayerY(index) = 0 Then Exit Sub 'Si esta al limite Y ignoramos

If GetPlayerY(index) = 1 Then


'Reciclamos CuboSup
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
b = buffer.ReadString
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
A = buffer.ReadLong
b = buffer.ReadString

'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString



Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True

ElseIf Map(mapa).Tile(Mapax, Mapay).ParteCubo = False Then

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False


'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString


Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True


'Cubo inferior 3 (Efecto difuso 2.5 D si se crea un inferior traspasable)

Map(mapa).Tile(Mapax2, Mapay2 + 1).Layer(Capa).Tileset = 1
Map(mapa).Tile(Mapax2, Mapay2 + 1).Layer(Capa).x = 6
Map(mapa).Tile(Mapax2, Mapay2 + 1).Layer(Capa).y = 6
Map(mapa).Tile(Mapax2, Mapay2 + 1).Type = TILE_TYPE_WALKABLE

Map(mapa).Tile(Mapax2, Mapay2 + 1).Data1 = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Data2 = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Data3 = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Data4 = ""
Map(mapa).Tile(Mapax2, Mapay2 + 1).Cubo = False 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2 + 1).HP = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Animacion = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Banco = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Evento = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).BancoLlave = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Script = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Timer = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).SFX1 = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).SFX2 = 0
Map(mapa).Tile(Mapax2, Mapay2 + 1).Mensaje = ""

End If

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong



Map(mapa).Tile(Mapax, Mapay).ParteCubo = False


'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString

Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True



Case DIR_DOWN

If GetPlayerY(index) = Map(mapa).MaxY Then Exit Sub

If GetPlayerY(index) = Map(mapa).MaxY - 1 Or Map(mapa).Tile(Mapax, Mapay + 1).Cubo = True Then

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False

GoTo IgnorarCubo2 'Si esta al limite Y ignoramos

End If

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False


'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString

Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True



Case DIR_RIGHT


If GetPlayerX(index) = Map(mapa).MaxX Then Exit Sub

If GetPlayerY(index) = Map(mapa).MaxY Or Map(mapa).Tile(Mapax, Mapay + 1).Cubo = True Then

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False

GoTo IgnorarCubo2 'Si esta al limite Y ignoramos

End If

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong



Map(mapa).Tile(Mapax, Mapay).ParteCubo = False


'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString


Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True




Case DIR_LEFT

If GetPlayerX(index) = 0 Then Exit Sub

If GetPlayerY(index) = Map(mapa).MaxY Or Map(mapa).Tile(Mapax, Mapay + 1).Cubo = True Then

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False

GoTo IgnorarCubo2 'Si esta al limite Y ignoramos

End If

'Tile
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).x = TileX
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).y = TileY
Map(mapa).Tile(Mapax, Mapay).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax, Mapay).Type = tipo
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax, Mapay).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Cubo = True 'Siempre sera un Cubo
Map(mapa).Tile(Mapax, Mapay).HP = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Script = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax, Mapay).Mensaje = buffer.ReadString
Map(mapa).Tile(Mapax, Mapay).Objeto = buffer.ReadLong


Map(mapa).Tile(Mapax, Mapay).ParteCubo = False


'CUBO 2

TileX2 = buffer.ReadLong
TileY2 = buffer.ReadLong
Capa2 = buffer.ReadLong
Mapax2 = buffer.ReadLong
Mapay2 = buffer.ReadLong
Tipo2 = buffer.ReadLong



Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).Tileset = TileN
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).x = TileX2
Map(mapa).Tile(Mapax2, Mapay2).Layer(Capa).y = TileY2
Map(mapa).Tile(Mapax2, Mapay2).Type = Tipo2
'Tile Transportador/Trampa
Map(mapa).Tile(Mapax2, Mapay2).Data1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data3 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Data4 = buffer.ReadString
Map(mapa).Tile(Mapax2, Mapay2).Cubo = True 'Siempre sera un Cubo

Map(mapa).Tile(Mapax2, Mapay2).HP = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Animacion = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Banco = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Evento = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).BancoLlave = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Script = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Timer = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX1 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).SFX2 = buffer.ReadLong
Map(mapa).Tile(Mapax2, Mapay2).Mensaje = buffer.ReadString

Map(mapa).Tile(Mapax2, Mapay2).ParteCubo = True



End Select

IgnorarCubo2: 'Se usa esta funcion para los limites Y del cubo inferior

    Call SpawnMapItems(mapa)


    ' Guardar y limpiar
    Call SaveMap(mapa)
     Call MapCache_Create(mapa)

    Call ClearTempTile(mapa)
    
    ' Enviar a todos
    For I = 1 To Player_HighIndex
        If isPlaying(I) And GetPlayerMap(I) = mapa Then
            Call PlayerWarp(I, mapa, GetPlayerX(I), GetPlayerY(I))
            'Call SendMapItemsToAll(I)
        End If
    Next I
    
 
     Call CacheMapBlocks(mapa)

    Set buffer = Nothing
    
'Inventario EaSee Engine 0.4 Version Cubos by ikicapitan (Podria utilizarse para sistema de flechas inclusive)
Equipamiento = GetPlayerEquipment(index, Weapon)
Inventario = 0

For I = 1 To MAX_INV 'Busca en el inventario
Inventario = GetPlayerInvItemNum(index, I)
Cantidad = GetPlayerInvItemValue(index, I)
If Inventario = Equipamiento Then 'Si hay un objeto que coincida con el objeto equipado
TakeInvItem index, Inventario, 1  ' Toma el objeto del inventario en cantidad 1

GoTo ActualizarEquipamiento
End If

Next

SetPlayerEquipment index, 0, Weapon 'Sino lo quita del equipado

ActualizarEquipamiento:

SendMapEquipment index
SendInventory index
   
End Sub



Sub MapaRespawnCubos(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    'Es una version alterada de handlemaprespawn sin actualizar NPC ni requerir privilegios
    
    
    CacheResources GetPlayerMap(index)
    
    Call AddLog(GetPlayerName(index) & " ha actualizado el mapa #" & GetPlayerMap(index), ADMIN_LOG)
End Sub

Sub Procesar_Golpe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim GolpeEmo As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    GolpeEmo = buffer.ReadLong
   
    Set buffer = Nothing
    
    Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) - GolpeEmo) 'Modificamos HP
    Call SendVital(index, Vitals.HP) 'Enviamos
    
End Sub

Sub Procesar_Visibilidad(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'EaSee 0.5

    Player(index).Visible = 0
    
    Call SendPlayerData(index)

End Sub


Sub Procesar_Buff(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim PointType As Byte
Dim buffer As clsBuffer
Dim BuffVal As Byte


    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte 'CLng(Parse(1))
    BuffVal = buffer.ReadByte
    Set buffer = Nothing
    
        ' make sure they're not maxed#
        If GetPlayerRawStat(index, PointType) >= 255 Then
            Exit Sub
        End If
        
        ' Everything is ok
    If BuffVal = 1 Then
    With Class(GetPlayerClass(index))
        Select Case PointType
            Case Stats.strength
                Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) - Player(index).FuerzaBuff)
                Player(index).FuerzaBuff = 0 'Fin del Buff
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) - Player(index).DestrezaBuff)
                Player(index).DestrezaBuff = 0 'Fin del Buff
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) - Player(index).InteligenciaBuff)
                Player(index).InteligenciaBuff = 0 'Fin del Buff
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) - Player(index).AgilidadBuff)
                Player(index).AgilidadBuff = 0 'Fin del Buff
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) - Player(index).VoluntadBuff)
                Player(index).VoluntadBuff = 0 'Fin del Buff
        End Select
        
        End With
        
    Else
    With Class(GetPlayerClass(index))
        Select Case PointType
            Case Stats.strength
                Call SetPlayerStat(index, Stats.strength, GetPlayerRawStat(index, Stats.strength) + Player(index).FuerzaDeBuff)
                Player(index).FuerzaDeBuff = 0 'Fin del DeBuff
            Case Stats.Endurance
                Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + Player(index).DestrezaDeBuff)
                Player(index).DestrezaDeBuff = 0 'Fin del DeBuff
            Case Stats.Intelligence
                Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + Player(index).InteligenciaDeBuff)
                Player(index).InteligenciaDeBuff = 0 'Fin del DeBuff
            Case Stats.Agility
                Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + Player(index).AgilidadDeBuff)
                Player(index).AgilidadDeBuff = 0 'Fin del DeBuff
            Case Stats.Willpower
                Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + Player(index).VoluntadDeBuff)
                Player(index).VoluntadDeBuff = 0 'Fin del DeBuff
        End Select
        
        End With
    
    
End If
    SendPlayerData index
End Sub

Sub Procesar_RestaurarSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
'EaSee 0.5

If Player(index).SpriteOR <> 0 And index <> 0 Then
    Player(index).Sprite = Player(index).SpriteOR

    Call SendPlayerData(index)
End If

End Sub

' ::::::::::::::::::::::::::
' :: Paquete Especial Para Cubos 0.9
' ::::::::::::::::::::::::::
Sub HandleCubo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Long
    Dim N As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim x As Long, y As Long
    
    ' can't attack whilst casting
    If TempPlayer(index).spellBuffer.Spell > 0 Then Exit Sub
    
    ' can't attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index

      ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
    End Select
    
    CheckCubos index, x, y 'EaSee 0.4 Cubos Ataque
End Sub
