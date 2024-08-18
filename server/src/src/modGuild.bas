Attribute VB_Name = "modGuild"
'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Max Members Per Guild
Public Const MAX_GUILD_MEMBERS As Long = 50
'Max Ranks Guilds Can Have
Public Const MAX_GUILD_RANKS As Long = 6
'Max Different Permissions
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6
'Max guild save files(aka max guilds)
Public Const MAX_GUILD_SAVES As Long = 200

'Default Ranks Info
'1: Open Admin
'2: Can Recruit
'3: Can Kick
'4: Can Edit Ranks
'5: Can Edit Users
'6: Can Edit Options

Public Guild_Ranks_Premission_Names(1 To MAX_GUILD_RANKS_PERMISSION) As String
Public Default_Ranks(1 To MAX_GUILD_RANKS_PERMISSION) As Byte


'Max is set to MAX_PLAYERS so each online player can have his own guild
Public GuildData(1 To MAX_PLAYERS) As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    Name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
End Type

Public Type GuildMemberRec
    'User login/name
    Used As Boolean
    
    User_Login As String
    User_Name As String
    Founder As Boolean
    
    Online As Boolean
    
    'Guild Variables
    Rank As Integer
    Comment As String * 100
     
End Type

Public Type GuildRec
    In_Use As Boolean
    
    Guild_Name As String
    Guild_Tag As String
    
    'Guild file number for saving
    Guild_Fileid As Long
    
    Guild_Members(1 To MAX_GUILD_MEMBERS) As GuildMemberRec
    Guild_Ranks(1 To MAX_GUILD_RANKS) As GuildRanksRec
    
    'Message of the day
    Guild_MOTD As String * 100
    
    'The rank recruits start at
    Guild_RecruitRank As Integer
    'Color of guild name
    Guild_Color As Integer

End Type

Public Sub Set_Default_Guild_Ranks()
    'Max sure this starts at 1 and ends at MAX_GUILD_RANKS_PERMISSION (Default 7)
    '0 = Cannot, 1 = Able To
    Guild_Ranks_Premission_Names(1) = "Administradores"
    Default_Ranks(1) = 0
    
    Guild_Ranks_Premission_Names(2) = "Puede Reclutar"
    Default_Ranks(2) = 1
    
    Guild_Ranks_Premission_Names(3) = "Puede Expulsar"
    Default_Ranks(3) = 0
    
    Guild_Ranks_Premission_Names(4) = "Puede Editar Rangos"
    Default_Ranks(4) = 0
    
    Guild_Ranks_Premission_Names(5) = "Puede Editar Usuarios"
    Default_Ranks(5) = 0
    
    Guild_Ranks_Premission_Names(6) = "Puede Editar Opciones"
    Default_Ranks(6) = 0
End Sub
Public Function GuildCheckName(index As Long, MemberSlot As Long, AttemptCorrect As Boolean) As Boolean
Dim I As Integer

    If Player(index).GuildFileId = 0 Or TempPlayer(index).tmpGuildSlot = 0 Or isPlaying(index) = False Or MemberSlot = 0 Then
        GuildCheckName = False
        Exit Function
    End If
    
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(MemberSlot).User_Login = Player(index).Login Then
        GuildCheckName = True
        Exit Function
    End If
    
    If AttemptCorrect = True Then
        If TempPlayer(index).tmpGuildSlot > 0 And Player(index).GuildFileId > 0 Then
            'did they get moved?
            For I = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).User_Login = Player(index).Login Then
                    Player(index).GuildMemberId = I
                    Call SavePlayer(index)
                    GuildCheckName = True
                    Exit Function
                Else
                    Player(index).GuildMemberId = 0
                End If
            Next I
                
            'Remove from guild if we can't find them
            If Player(index).GuildMemberId = 0 Then
                Player(index).GuildFileId = 0
                TempPlayer(index).tmpGuildSlot = 0
                Call SavePlayer(index)
                PlayerMsg index, "No podemos encontrarte en tu lista de Clanes por 2 posibles motivos.", BrightRed
                PlayerMsg index, "1)Has sido expulsado   2)Tu Clan ha sido eliminado o reemplazado", BrightRed
            End If
        End If
    End If
    
    
    GuildCheckName = False


End Function
Public Sub MakeGuild(Founder_Index As Long, Name As String, Tag As String)
    Dim tmpGuild As GuildRec
    Dim GuildSlot As Long
    Dim GuildFileId As Long
    Dim I As Integer
    Dim b As Integer
    Dim itemAmount As Long
    
    If Player(Founder_Index).GuildFileId > 0 Then
        PlayerMsg Founder_Index, "Primero debes abandonar el Clan!", BrightRed
        Exit Sub
    End If
    
    GuildFileId = Find_Guild_Save
    GuildSlot = FindOpenGuildSlot
    
    If Not isPlaying(Founder_Index) Then Exit Sub
    
    'We are unable for an unknown reason
    If GuildSlot = 0 Or GuildFileId = 0 Then
        PlayerMsg Founder_Index, "Imposible crear Clan!", BrightRed
        Exit Sub
    End If
    
    If Name = "" Then
        PlayerMsg Founder_Index, "Tu Clan necesita un nombre!", BrightRed
        Exit Sub
    End If
    
    ' Check level
    If GetPlayerLevel(Founder_Index) < Options.Buy_Lvl Then
        PlayerMsg Founder_Index, "Requiere nivel " & Options.Buy_Lvl & " para crear un Clan!", BrightRed
        Exit Sub
    End If
    
    ' Check if item is required
    If Not Options.Buy_Item = 0 Then
        'Get item amount
        itemAmount = HasItem(Founder_Index, Options.Buy_Item)
                    
        ' Item Req
        If itemAmount = 0 Or itemAmount < Options.Buy_Cost Then
            PlayerMsg Founder_Index, "Necesitas " & Options.Buy_Cost & " " & Item(Options.Buy_Item).Name & " para unirte a un Clan!", BrightRed
            Exit Sub
        End If
                
        'Take Item
        TakeInvItem Founder_Index, Options.Buy_Item, Options.Buy_Cost
    End If
    
    GuildData(GuildSlot).Guild_Name = Name
    GuildData(GuildSlot).Guild_Tag = Tag
    GuildData(GuildSlot).Guild_Color = 4
    GuildData(GuildSlot).Guild_MOTD = "Bienvenido a " & Name & "!"
    GuildData(GuildSlot).In_Use = True
    GuildData(GuildSlot).Guild_Fileid = GuildFileId
    GuildData(GuildSlot).Guild_Members(1).Founder = True
    GuildData(GuildSlot).Guild_Members(1).User_Login = Player(Founder_Index).Login
    GuildData(GuildSlot).Guild_Members(1).User_Name = Player(Founder_Index).Name
    GuildData(GuildSlot).Guild_Members(1).Rank = MAX_GUILD_RANKS
    GuildData(GuildSlot).Guild_Members(1).Comment = "Fundador"
    GuildData(GuildSlot).Guild_Members(1).Used = True
    GuildData(GuildSlot).Guild_Members(1).Online = True
    

    'Set up Admin Rank with all permission which is just the max rank
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Name = "Lider"
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Used = True
    
    For b = 1 To MAX_GUILD_RANKS_PERMISSION
        GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).RankPermission(b) = 1
    Next b
    
    'Set up rest of the ranks with default permission
    For I = 1 To MAX_GUILD_RANKS - 1
        GuildData(GuildSlot).Guild_Ranks(I).Name = "Rango " & I
        GuildData(GuildSlot).Guild_Ranks(I).Used = True
        
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData(GuildSlot).Guild_Ranks(I).RankPermission(b) = Default_Ranks(b)
        Next b

    Next I
    
    Player(Founder_Index).GuildFileId = GuildFileId
    Player(Founder_Index).GuildMemberId = 1
    TempPlayer(Founder_Index).tmpGuildSlot = GuildSlot
    
    
    'Save
    Call SaveGuild(GuildSlot)
    Call SavePlayer(Founder_Index)
    
    'Send to player
    Call SendGuild(False, Founder_Index, GuildSlot)
    
    'Inform users
    PlayerMsg Founder_Index, "Clan creado", BrightGreen
    PlayerMsg Founder_Index, "Bienvenido a " & GuildData(GuildSlot).Guild_Name & ".", BrightGreen
    
    PlayerMsg Founder_Index, "Para conversar con el Clan usa en el Chat:  ;mensaje ", BrightRed
    
    'Update user for guild name display
    Call SendPlayerData(Founder_Index)

    
End Sub
Public Function CheckGuildPermission(index As Long, Permission As Integer) As Boolean
Dim GuildSlot As Long

    'Get slot
    GuildSlot = TempPlayer(index).tmpGuildSlot
    
    'Make sure we are talking about the same person
    If Not GuildData(GuildSlot).Guild_Members(Player(index).GuildMemberId).User_Login = Player(index).Login Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'If founder, true in every case
    If GuildData(GuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = True Then
        CheckGuildPermission = True
        Exit Function
    End If
    
    'Make sure this slot is being used aka they are still a member
    If GuildData(GuildSlot).Guild_Members(Player(index).GuildMemberId).Used = False Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'Check if they are able to
    If GuildData(GuildSlot).Guild_Ranks(GuildData(GuildSlot).Guild_Members(Player(index).GuildMemberId).Rank).RankPermission(Permission) = 1 Then
        CheckGuildPermission = True
    Else
        CheckGuildPermission = False
    End If
    
End Function
Public Sub Request_Guild_Invite(index As Long, GuildSlot As Long, Inviter_Index As Long)

    If Player(index).GuildFileId > 0 Then
        PlayerMsg index, "Debes abandonar tu Clan para unirte a " & GuildData(GuildSlot).Guild_Name & "!", BrightRed
        PlayerMsg Inviter_Index, "Ya pertenecen a un Clan!", BrightRed
        Exit Sub
    End If

    If TempPlayer(index).tmpGuildInviteSlot > 0 Then
        PlayerMsg Inviter_Index, "El usuario tiene una invitacion pendiente. Intenta nuevamente.", BrightRed
        Exit Sub
    End If

    'Permission 2 = Can Recruit
    If CheckGuildPermission(Inviter_Index, 2) = False Then
        PlayerMsg Inviter_Index, "Rango insuficiente!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(index).tmpGuildInviteSlot = GuildSlot
    '2 minute
    TempPlayer(index).tmpGuildInviteTimer = GetTickCount + 120000
    
    TempPlayer(index).tmpGuildInviteId = Player(Inviter_Index).GuildFileId
    
    PlayerMsg Inviter_Index, "Invitacion enviada!", Green
    PlayerMsg index, Trim$(Player(Inviter_Index).Name) & " te ha invitado a unirte a su Clan " & GuildData(GuildSlot).Guild_Name & "!", Green
    PlayerMsg index, "Escribe  /guild accept   en los 2 proximos minutos para unirte.", Green
    PlayerMsg index, "Escribe  /guild decline   para rechazar la peticion.", Green
End Sub

Public Sub Join_Guild(index As Long, GuildSlot As Long)
Dim OpenSlot As Long
    
    If isPlaying(index) = False Then Exit Sub
    
    OpenSlot = FindOpenGuildMemberSlot(GuildSlot)
        'Guild full?
        If OpenSlot > 0 Then
        
            ' Check level
            If GetPlayerLevel(index) < Options.Join_Lvl Then
                PlayerMsg index, "Necesitas ser nivel " & Options.Join_Lvl & " para unirte a un Clan!", BrightRed
                Exit Sub
            End If
            
            ' Check if item is required
            If Not Options.Join_Item = 0 Then
                'Get item amount
                itemAmount = HasItem(index, Options.Join_Item)
                    
                ' Gold Req
                If itemAmount = 0 Or itemAmount < Options.Join_Cost Then
                    PlayerMsg index, "Necesitas " & Options.Join_Cost & " " & Item(Options.Join_Item).Name & " para unirte a un Clan!", BrightRed
                    Exit Sub
                End If
                
                'Take Item
                TakeInvItem index, Options.Join_Item, Options.Join_Cost
            End If
        
            'Set guild data
            GuildData(GuildSlot).Guild_Members(OpenSlot).Used = True
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Login = Player(index).Login
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Name = Player(index).Name
            GuildData(GuildSlot).Guild_Members(OpenSlot).Rank = GuildData(GuildSlot).Guild_RecruitRank
            GuildData(GuildSlot).Guild_Members(OpenSlot).Comment = "Miembro: " & DateValue(Now)
            GuildData(GuildSlot).Guild_Members(OpenSlot).Online = True
            
            'Set player data
            Player(index).GuildFileId = GuildData(GuildSlot).Guild_Fileid
            Player(index).GuildMemberId = OpenSlot
            TempPlayer(index).tmpGuildSlot = GuildSlot
            
            'Save
            Call SaveGuild(GuildSlot)
            Call SavePlayer(index)
            
            'Send player guild data and display welcome
            Call SendGuild(True, index, GuildSlot)
            PlayerMsg index, "Bienvenido a " & GuildData(GuildSlot).Guild_Name & ".", BrightGreen
            
            PlayerMsg index, "Puedes conversar con tu Clan escribiendo:  ;mensaje", BrightGreen
            
            'Update player to display guild name
            Call SendPlayerData(index)
            
        Else
            'Guild full display msg
            PlayerMsg index, "Clan lleno", BrightRed
        End If
    
End Sub

Public Function Find_Guild_Save() As Long
Dim FoundSlot As Boolean
Dim Current As Integer
FoundSlot = False
Current = 1

Do Until FoundSlot = True
    
    If Not FileExist("\Data\guilds\Guild" & Current & ".dat") Then
        Find_Guild_Save = Current
        FoundSlot = True
    Else
        Current = Current + 1
    End If
    
    'Max Guild Files check
    If Current > MAX_GUILD_SAVES Then
        'send back 0 for no slot found
        Find_Guild_Save = 0
        FoundSlot = True
    End If
    
    
Loop

End Function
Public Function FindOpenGuildSlot() As Long
    Dim I As Integer
    
    For I = 1 To MAX_PLAYERS
        If GuildData(I).In_Use = False Then
            FindOpenGuildSlot = I
            Exit Function
        End If
        
        'No slot found how?
        FindOpenGuildSlot = 0
    Next I
End Function
Public Function FindOpenGuildMemberSlot(GuildSlot As Long) As Long
Dim I As Integer
    
    For I = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(I).Used = False Then
            FindOpenGuildMemberSlot = I
            Exit Function
        End If
    Next I
    
    'Guild is full sorry bub
    FindOpenGuildMemberSlot = 0

End Function
Public Sub ClearGuildMemberSlot(GuildSlot As Long, MembersSlot As Long)
    GuildData(GuildSlot).Guild_Members(MembersSlot).Used = False
    GuildData(GuildSlot).Guild_Members(MembersSlot).User_Login = vbNullString
    GuildData(GuildSlot).Guild_Members(MembersSlot).User_Name = vbNullString
    GuildData(GuildSlot).Guild_Members(MembersSlot).Rank = 0
    GuildData(GuildSlot).Guild_Members(MembersSlot).Comment = vbNullString
    GuildData(GuildSlot).Guild_Members(MembersSlot).Founder = False
    GuildData(GuildSlot).Guild_Members(MembersSlot).Online = False
            
    'Save guild after we remove member
    Call SaveGuild(GuildSlot)
End Sub
Public Sub LoadGuild(GuildSlot As Long, GuildFileId As Long)
Dim I As Integer
'If 0 something is wrong
If GuildFileId = 0 Then Exit Sub

    'Does this file even exist?
    If Not FileExist("\Data\guilds\Guild" & GuildFileId & ".dat") Then Exit Sub

    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\guilds\Guild" & GuildFileId & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , GuildData(GuildSlot)
    Close #F
        
    GuildData(GuildSlot).In_Use = True
        
    'Make sure an online flag didn't manage to slip through
    For I = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(I).Online = True Then
            GuildData(GuildSlot).Guild_Members(I).Online = False
        End If
    Next I
        
End Sub
Public Sub SaveGuild(GuildSlot As Long)

    'Dont save unless a fileid was assigned
    If GuildData(GuildSlot).Guild_Fileid = 0 Then Exit Sub


    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\guilds\Guild" & GuildData(GuildSlot).Guild_Fileid & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , GuildData(GuildSlot)
    Close #F
    
End Sub
Public Sub UnloadGuildSlot(GuildSlot As Long)
    'Exit on error
    If GuildSlot = 0 Or GuildSlot > MAX_GUILD_SAVES Then Exit Sub
    If GuildData(GuildSlot).In_Use = False Then Exit Sub
    
    'Save it first
    Call SaveGuild(GuildSlot)
    'Clear and reset for next use
    Call ClearGuild(GuildSlot)
End Sub
Public Sub ClearGuilds()
Dim I As Long

    For I = 1 To MAX_PLAYERS
        Call ClearGuild(I)
    Next I
End Sub
Public Sub ClearGuild(index As Long)
    Call ZeroMemory(ByVal VarPtr(GuildData(index)), LenB(GuildData(index)))
    GuildData(index).Guild_Name = vbNullString
    GuildData(index).Guild_Tag = vbNullString
    GuildData(index).In_Use = False
    GuildData(index).Guild_Fileid = 0
    GuildData(index).Guild_Color = 0
    GuildData(index).Guild_RecruitRank = 1
End Sub
Public Sub CheckUnloadGuild(GuildSlot As Long)
Dim I As Integer
Dim UnloadGuild As Boolean

UnloadGuild = True

If GuildData(GuildSlot).In_Use = False Then Exit Sub

    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            If Player(I).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                UnloadGuild = False
                Exit For
            End If
        End If
    Next I
    
    If UnloadGuild = True Then
        Call UnloadGuildSlot(GuildSlot)
    End If
End Sub
Public Sub GuildKick(GuildSlot As Long, index As Long, playerName As String)
Dim FoundOffline As Boolean
Dim IsOnline As Boolean
Dim OnlineIndex As Long
Dim MemberSlot As Long
Dim I As Integer
    
    OnlineIndex = FindPlayer(playerName)
    
    If OnlineIndex = index Then
        PlayerMsg index, "No puedes expulsarte!", BrightRed
        Exit Sub
    End If
    
    'If OnlineIndex > 0 they are online
    If OnlineIndex > 0 Then
        IsOnline = True
        
        If Player(OnlineIndex).GuildMemberId > 0 Then
            MemberSlot = Player(OnlineIndex).GuildMemberId
        Else
            'Prevent error, rest of this code assumes this is greater than 0
            Exit Sub
        End If
        
    Else
        IsOnline = False
    End If
    
    
    'Handle kicking online user
    If IsOnline = True Then
        If Not Player(index).GuildFileId = Player(OnlineIndex).GuildFileId Then
            PlayerMsg index, "No pertenece a tu Clan!", BrightRed
            Exit Sub
        End If
        
        If GuildData(GuildSlot).Guild_Members(MemberSlot).Founder = True Then
            PlayerMsg index, "No puedes expulsar al fundador!!", BrightRed
            Exit Sub
        End If
        
        Player(OnlineIndex).GuildFileId = 0
        Player(OnlineIndex).GuildMemberId = 0
        TempPlayer(OnlineIndex).tmpGuildSlot = 0
        Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
        PlayerMsg OnlineIndex, "Has sido expulsado del Clan!", BrightRed
        PlayerMsg index, "Jugador expulsado!", BrightRed
        Call SavePlayer(OnlineIndex)
        Call SaveGuild(GuildSlot)
        Call SendGuild(True, OnlineIndex, GuildSlot)
        Call SendGuild(True, index, GuildSlot)
        Call SendPlayerData(OnlineIndex)
        Exit Sub
    End If
    
    
    
    'Handle Kicking Offline User
    FoundOffline = False
    If IsOnline = False Then
        'Lets Try to find them in the roster
        For I = 1 To MAX_GUILD_MEMBERS
            If playerName = Trim$(GuildData(GuildSlot).Guild_Members(I).User_Name) Then
                'Found them
                FoundOffline = True
                MemberSlot = I
                Exit For
            End If
        Next I
        
        If FoundOffline = True Then
        
            If MemberSlot = 0 Then Exit Sub
            
            Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
            Call SaveGuild(GuildSlot)
            PlayerMsg index, "Jugador fuera de linea expulsado!", BrightRed
            Exit Sub
        End If
        
        If FoundOffline = False And IsOnline = False Then
            PlayerMsg index, "No se encuentra " & playerName & " en tu Clan.", BrightRed
        End If
    
    End If
 
End Sub
Public Sub GuildLeave(index As Long)
Dim I As Integer
Dim GuildSlot As Long

    
    'This is for the leave command only, kicking has its own sub because it handles both online and offline kicks, while this only handles online.
    
    If Not Player(index).GuildFileId > 0 Then
        PlayerMsg index, "No te encuentras en un Clan!", BrightRed
        Exit Sub
    End If
    
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = True Then
        PlayerMsg index, "Primero debes transferir el titulo de Fundador.", BrightRed
        PlayerMsg index, "Escribe /founder (nombre) para transferir el titulo.", BrightRed
        Exit Sub
    End If
    
    'They match so they can leave
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).User_Login = Player(index).Login Then
        
        GuildSlot = TempPlayer(index).tmpGuildSlot
        
        'Clear guild slot
        Call ClearGuildMemberSlot(TempPlayer(index).tmpGuildSlot, Player(index).GuildMemberId)
        
        'Clear player data
        Player(index).GuildFileId = 0
        Player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
        
            Player(index).GuildMemberId = OpenSlot
            TempPlayer(index).tmpGuildSlot = GuildSlot
        
        'Update user for guild name display
        Call SendGuild(True, OnlineIndex, GuildSlot)
        Call SendGuild(True, index, GuildSlot)
        Call SendPlayerData(index)
        
        PlayerMsg index, "Has abandonado el Clan.", BrightRed

    Else
        'They don't match this slot remove them
        Player(index).GuildFileId = 0
        Player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
    End If
    
    
End Sub
Public Sub GuildLoginCheck(index As Long)
Dim I As Long
Dim GuildSlot As Long
Dim GuildLoaded As Boolean
GuildLoaded = False


    'Not in guild
    If Player(index).GuildFileId = 0 Then Exit Sub
    
    'Check to make sure the guild file exists
    If Not FileExist("\Data\guilds\Guild" & Player(index).GuildFileId & ".dat") Then
        'If guild was deleted remove user from guild
        Player(index).GuildFileId = 0
        Player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
        Call SavePlayer(index)
        PlayerMsg index, "Tu Clan ha sido eliminado!", BrightRed
        Exit Sub
    End If
    
    'First we need to see if our guild is loaded
    For I = 1 To MAX_PLAYERS
        If GuildData(I).In_Use = True Then
            'If its already loaded set true
            If GuildData(I).Guild_Fileid = Player(index).GuildFileId Then
                GuildLoaded = True
                GuildSlot = I
                Exit For
            End If
        End If
    Next I
    
    'If the guild is not loaded we need to load it
    If GuildLoaded = False Then
        'Find open guild slot, if 0 none
        GuildSlot = FindOpenGuildSlot
        If GuildSlot > 0 Then
            'LoadGuild
            Call LoadGuild(GuildSlot, Player(index).GuildFileId)
            
        End If
    End If
    
    'Set GuildSlot
    TempPlayer(index).tmpGuildSlot = GuildSlot
    
    'This is to prevent errors when we look for them
    If Player(index).GuildMemberId = 0 Then Player(index).GuildMemberId = 1

    'Make sure user didn't get kicked or guild was replaced by a different guild, both result in removal
    If GuildCheckName(index, Player(index).GuildMemberId, True) = False Then
        'unload if this user is not in this guild and it was loaded for this user
        If GuildLoaded = False Then
            Call UnloadGuildSlot(GuildSlot)
            Exit Sub
        End If
    End If
    
    'Sent data and set slot if all is good
    If Player(index).GuildFileId > 0 Then
        'Set online flag
        GuildData(GuildSlot).Guild_Members(Player(index).GuildMemberId).Online = True
        
        
        'send
        Call SendGuild(False, index, GuildSlot)
        
        'Display motd
        If Not GuildData(GuildSlot).Guild_MOTD = vbNullString Then
            PlayerMsg index, "Clan Mensaje: " & GuildData(GuildSlot).Guild_MOTD, Blue
        End If
    End If
    
End Sub
Sub DisbandGuild(GuildSlot As Long, index As Long)
Dim I As Integer
Dim tmpGuildSlot As Long
Dim TmpGuildFileId As Long
Dim filename As String

'Set some thing we need
tmpGuildSlot = GuildSlot
TmpGuildFileId = GuildData(tmpGuildSlot).Guild_Fileid

    'They are who they say they are, and are founder
    If GuildCheckName(index, Player(index).GuildMemberId, False) = True And GuildData(tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = True Then
        'File exists right?
         If FileExist("\Data\Guilds\Guild" & TmpGuildFileId & ".dat") = True Then
            'We have a go for disband
            'First we take everyone online out, this will include the founder people who login later will be kicked out then
            For I = 1 To Player_HighIndex
                If isPlaying(I) = True Then
                    If Player(I).GuildFileId = TmpGuildFileId Then
                        'remove from guild
                        Player(I).GuildFileId = 0
                        Player(I).GuildMemberId = 0
                        TempPlayer(I).tmpGuildSlot = 0
                        Call SavePlayer(I)
                        'Send player data so they don't have name over head anymore
                        Call SendPlayerData(I)
                    End If
                End If
            Next I
            
            'Unload Guild from memory
            Call UnloadGuildSlot(tmpGuildSlot)

            filename = App.Path & "\Data\Guilds\Guild" & TmpGuildFileId & ".dat"
            Kill filename
            
            
            PlayerMsg index, "Clan disuelto!", BrightGreen
         End If
    Else
        PlayerMsg index, "No puedes hacer eso!", BrightRed
    End If
End Sub
Sub SendDataToGuild(ByVal GuildSlot As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If isPlaying(I) Then
            If Player(I).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next

End Sub

Sub SendGuild(ByVal SendToWholeGuild As Boolean, ByVal index As Long, ByVal GuildSlot)
    Dim buffer As clsBuffer
    Dim I As Integer
    Dim b As Integer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SSendGuild
    
    'General data
    buffer.WriteString GuildData(GuildSlot).Guild_Name
    buffer.WriteString GuildData(GuildSlot).Guild_Tag
    buffer.WriteInteger GuildData(GuildSlot).Guild_Color
    buffer.WriteString GuildData(GuildSlot).Guild_MOTD
    buffer.WriteInteger GuildData(GuildSlot).Guild_RecruitRank
    
    
    'Send Members
    For I = 1 To MAX_GUILD_MEMBERS
        buffer.WriteString GuildData(GuildSlot).Guild_Members(I).User_Name
        buffer.WriteInteger GuildData(GuildSlot).Guild_Members(I).Rank
        buffer.WriteString GuildData(GuildSlot).Guild_Members(I).Comment
        buffer.WriteByte GuildData(GuildSlot).Guild_Members(I).Online
    Next I
    
    'Send Ranks
    For I = 1 To MAX_GUILD_RANKS
            buffer.WriteString GuildData(GuildSlot).Guild_Ranks(I).Name
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            buffer.WriteByte GuildData(GuildSlot).Guild_Ranks(I).RankPermission(b)
            buffer.WriteString Guild_Ranks_Premission_Names(b)
        Next b
    Next I
    
    If SendToWholeGuild = False Then
        SendDataTo index, buffer.ToArray()
    Else
        SendDataToGuild GuildSlot, buffer.ToArray()
    End If
    
    For I = 1 To MAX_GUILD_MEMBERS
        SendPlayerData I
    Next I
    
    Set buffer = Nothing
End Sub
Sub ToggleGuildAdmin(ByVal index As Long, ByVal GuildSlot, ByVal OpenAdmin As Boolean)
    Dim buffer As clsBuffer
    Dim I As Integer
    Dim b As Integer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SAdminGuild
    
    
    If OpenAdmin = True Then
        buffer.WriteByte 1
    Else
        buffer.WriteByte 0
    End If

        SendDataTo index, buffer.ToArray()

    
    Set buffer = Nothing
End Sub
Sub SayMsg_Guild(ByVal GuildSlot As Long, ByVal index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString message
    buffer.WriteString "[" & GuildData(GuildSlot).Guild_Tag & "]"
    buffer.WriteLong saycolour
    
    SendDataToGuild GuildSlot, buffer.ToArray()

    
    Set buffer = Nothing
End Sub
Public Sub HandleGuildMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
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
    
    If Not Player(index).GuildFileId > 0 Then
        PlayerMsg index, "No perteneces a ningun Clan!", BrightRed
        Exit Sub
    End If
    
    s = "[" & GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag & "]" & GetPlayerName(index) & ": " & Msg
    
    Call SayMsg_Guild(TempPlayer(index).tmpGuildSlot, index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(Msg)
    
    Set buffer = Nothing
End Sub
Public Sub HandleGuildSave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim buffer As clsBuffer
Dim SaveType As Integer
Dim SentIndex As Integer
Dim HoldInt As Integer
Dim I As Integer


    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    
    SaveType = buffer.ReadInteger
    SentIndex = buffer.ReadInteger
    
    If SaveType = 0 Or SentIndex = 0 Then Exit Sub
    
    
    Select Case SaveType
    Case 1
        'options
        If CheckGuildPermission(index, 6) = True Then
            
            'Guild Color
            HoldInt = buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(index).tmpGuildSlot).Guild_Color = HoldInt
                HoldInt = 0
            End If
            
            'Guild Recruit rank
            HoldInt = buffer.ReadInteger
            
            'Guild Name
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name = buffer.ReadString
            
            'Guild Tag
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Tag = buffer.ReadString
            
            'Guild MOTD
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_MOTD = buffer.ReadString
            
            
            'Did Recruit Rank change? Make sure they didnt set recruit rank at or above their rank
            If Not GuildData(TempPlayer(index).tmpGuildSlot).Guild_RecruitRank = HoldInt Then
                If Not HoldInt >= GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Rank Then
                    GuildData(TempPlayer(index).tmpGuildSlot).Guild_RecruitRank = HoldInt
                    
                Else
                    PlayerMsg index, "No deberias permitir que el rango reclutador sea de mayor o igual jerarquia al tuyo.", BrightRed
                End If
            End If
        Else
            PlayerMsg index, "No permitido.", BrightRed
        End If
        HoldInt = 0
    Case 2
        'users
        If CheckGuildPermission(index, 5) = True Then
            'Guild Member Rank
            HoldInt = buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Rank = HoldInt
            Else
                PlayerMsg index, "Debes modificar el rango sobre 0", BrightRed
            End If
            
            'Guild Member Comment
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Comment = buffer.ReadString
        Else
            PlayerMsg index, "No permitido.", BrightRed
        End If
        
    Case 3
        'ranks
        If CheckGuildPermission(index, 4) = True Then
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).Name = buffer.ReadString
                For I = 1 To MAX_GUILD_RANKS_PERMISSION
                    GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).RankPermission(I) = buffer.ReadByte
                Next I
        Else
            PlayerMsg index, "No permitido.", BrightRed
        End If
    
    End Select
    
    Call SendGuild(True, index, TempPlayer(index).tmpGuildSlot)
    
    Set buffer = Nothing
End Sub
Public Sub HandleGuildCommands(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim I As Integer
    Dim SelectedIndex As Long
    Dim SendText As String
    Dim SendText2 As String
    Dim SelectedCommand As Integer
    Dim MembersCount As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    
    SelectedCommand = buffer.ReadInteger
    SendText = buffer.ReadString
    
    If SelectedCommand = 1 Then
        SendText2 = buffer.ReadString
    End If
    
    'Command 1/6/7 can be used while not in a guild
    If Player(index).GuildFileId = 0 And Not (SelectedCommand = 1 Or SelectedCommand = 6 Or SelectedCommand = 7) Then
        PlayerMsg index, "No perteneces a un Clan!", BrightRed
        Exit Sub
    End If
    
    Select Case SelectedCommand
    Case 1
        'make
        Call MakeGuild(index, SendText, SendText2)
        PlayerMsg index, SendText & " - " & SendText2, BrightRed
        
    Case 2
        'invite
        'Find user index
        SelectedIndex = 0
        
        'Try to find player
        SelectedIndex = FindPlayer(SendText)
        
        If SelectedIndex > 0 Then
            Call Request_Guild_Invite(SelectedIndex, TempPlayer(index).tmpGuildSlot, index)
        Else
            PlayerMsg index, "No se encuentra el usuario " & SendText & ".", BrightRed
        End If
        
    Case 3
        'leave
        Call GuildLeave(index)
        
    Case 4
        'admin
        If CheckGuildPermission(index, 1) = True Then
            Call ToggleGuildAdmin(index, TempPlayer(index).tmpGuildSlot, True)
        Else
            PlayerMsg index, "No permitido.", BrightRed
        End If
    
    Case 5
        'view
        'This sets the default option
        If SendText = "" Then SendText = "en linea"
        MembersCount = 0
        
        Select Case SendText
        Case "en linea"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For I = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).Online = True Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next I
            
            PlayerMsg index, "Total: " & MembersCount, Green
        
        Case "todos"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For I = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).Used = True Then
                    PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).User_Name, Green
                    MembersCount = MembersCount + 1
                End If
            Next I
            
            PlayerMsg index, "Total: " & MembersCount, Green
        
        Case "fuera de linea"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For I = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).Online = False Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(I).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next I
            
            PlayerMsg index, "Total: " & MembersCount, Green
        
        End Select
    Case 6
        'accept
        If TempPlayer(index).tmpGuildInviteSlot > 0 Then
            If GuildData(TempPlayer(index).tmpGuildInviteSlot).In_Use = True And GuildData(TempPlayer(index).tmpGuildInviteSlot).Guild_Fileid = TempPlayer(index).tmpGuildInviteId Then
                Call Join_Guild(index, TempPlayer(index).tmpGuildInviteSlot)
                TempPlayer(index).tmpGuildInviteSlot = 0
                TempPlayer(index).tmpGuildInviteTimer = 0
                TempPlayer(index).tmpGuildInviteId = 0
            Else
                PlayerMsg index, "Nadie de este Clan se encuentra en linea.", BrightRed
            End If
        Else
            PlayerMsg index, "Debes tener una solicitud de Clan.", BrightRed
        End If
    Case 7
        'decline
        If TempPlayer(index).tmpGuildInviteSlot > 0 Then
            TempPlayer(index).tmpGuildInviteSlot = 0
            TempPlayer(index).tmpGuildInviteTimer = 0
            TempPlayer(index).tmpGuildInviteId = 0
            PlayerMsg index, "Has rechazado la solicitud.", BrightRed
        Else
            PlayerMsg index, "Debes tener una solicitud de Clan.", BrightRed
        End If
        
    Case 8
        'founder
        'Make sure the person who used the command is who they say they are
        If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).User_Login = Player(index).Login Then
            'Make sure they are founder
            If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = True Then
                'Find user index
                SelectedIndex = 0
                
                'Try to find player
                SelectedIndex = FindPlayer(SendText)
                
                If SelectedIndex > 0 Then
                    'Make sure the person getting founder is the correct person
                    If GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(Player(SelectedIndex).GuildMemberId).User_Login = Player(SelectedIndex).Login Then
                        GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = False
                        GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(Player(SelectedIndex).GuildMemberId).Founder = True
                    End If
                Else
                    PlayerMsg index, "No puedes encontrar al usuario " & SendText & ".", BrightRed
                End If
            Else
                 PlayerMsg index, "Solo el Fundador puede realizar esta accion.", BrightRed
            End If
        End If
    Case 9
        'kick
        Call GuildKick(TempPlayer(index).tmpGuildSlot, index, SendText)
    
    Case 10
        'disband
        Call DisbandGuild(TempPlayer(index).tmpGuildSlot, index)
        
    
    End Select
  
    Set buffer = Nothing
End Sub
