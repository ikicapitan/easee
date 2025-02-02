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
Dim i As Integer
Dim Var1 As String

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
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Login = Player(index).Login Then
                    Player(index).GuildMemberId = i
                    Call SavePlayer(index)
                    GuildCheckName = True
                    Exit Function
                Else
                    Player(index).GuildMemberId = 0
                End If
            Next i
                
            'Remove from guild if we can't find them
            If Player(index).GuildMemberId = 0 Then
                Player(index).GuildFileId = 0
                TempPlayer(index).tmpGuildSlot = 0
                Call SavePlayer(index)
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M1")
                PlayerMsg index, Var1, BrightRed
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M2")
                PlayerMsg index, Var1, BrightRed
                                    '[GREMIOS]M1 y M2
            End If
        End If
    End If
    
    
    GuildCheckName = False


End Function
Public Sub MakeGuild(Founder_Index As Long, Name As String, Tag As String)
    Dim tmpGuild As GuildRec
    Dim GuildSlot As Long
    Dim GuildFileId As Long
    Dim i As Integer
    Dim b As Integer
    Dim itemAmount As Long
    Dim Var1 As String
    
    If Player(Founder_Index).GuildFileId > 0 Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M3")
        PlayerMsg Founder_Index, Var1, BrightRed 'Primero debes abandonar el Clan!
        Exit Sub
    End If
    
    GuildFileId = Find_Guild_Save
    GuildSlot = FindOpenGuildSlot
    
    If Not isPlaying(Founder_Index) Then Exit Sub
    
    'We are unable for an unknown reason
    If GuildSlot = 0 Or GuildFileId = 0 Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M4")
        PlayerMsg Founder_Index, Var1, BrightRed 'Imposible crear Clan!
        Exit Sub
    End If
    
    If Name = "" Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M5")
        PlayerMsg Founder_Index, Var1, BrightRed 'Tu Clan necesita un nombre!
        Exit Sub
    End If
    
    ' Check level
    If GetPlayerLevel(Founder_Index) < Options.Buy_Lvl Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M6")
        PlayerMsg Founder_Index, Var1 & Options.Buy_Lvl & "!", BrightRed
        Exit Sub                        'Para crear un clan necesitas nivel
    End If
    
    ' Check if item is required
    If Not Options.Buy_Item = 0 Then
        'Get item amount
        itemAmount = HasItem(Founder_Index, Options.Buy_Item)
                    
        ' Item Req
        If itemAmount = 0 Or itemAmount < Options.Buy_Cost Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M7")
            PlayerMsg Founder_Index, Var1 & Options.Buy_Cost & " " & Item(Options.Buy_Item).Name & "!", BrightRed
            Exit Sub                '[GREMIOS]M7
        End If
                
        'Take Item
        TakeInvItem Founder_Index, Options.Buy_Item, Options.Buy_Cost
    End If
    
    GuildData(GuildSlot).Guild_Name = Name
    GuildData(GuildSlot).Guild_Tag = Tag
    GuildData(GuildSlot).Guild_Color = 4
    GuildData(GuildSlot).Guild_MOTD = "Bienvenido a " & Name & "."
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
    For i = 1 To MAX_GUILD_RANKS - 1
        GuildData(GuildSlot).Guild_Ranks(i).Name = "Rango " & i
        GuildData(GuildSlot).Guild_Ranks(i).Used = True
        
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b) = Default_Ranks(b)
        Next b

    Next i
    
    Player(Founder_Index).GuildFileId = GuildFileId
    Player(Founder_Index).GuildMemberId = 1
    TempPlayer(Founder_Index).tmpGuildSlot = GuildSlot
    
    
    'Save
    Call SaveGuild(GuildSlot)
    Call SavePlayer(Founder_Index)
    
    'Send to player
    Call SendGuild(False, Founder_Index, GuildSlot)
    
    'Inform users
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M8")
    PlayerMsg Founder_Index, Var1, BrightGreen '[GREMIOS]M8
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M9")
    PlayerMsg Founder_Index, Var1 & GuildData(GuildSlot).Guild_Name & ".", BrightGreen '[GREMIOS]M9
    
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M10")
    PlayerMsg Founder_Index, Var1, BrightRed '[GREMIOS]M10
    
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
Dim Var1 As String

    If Player(index).GuildFileId > 0 Then '[GREMIOS]M11
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M11")
        PlayerMsg index, Var1 & GuildData(GuildSlot).Guild_Name & ".", BrightRed
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M12")
        PlayerMsg Inviter_Index, Var1, BrightRed
        Exit Sub                        '[GREMIOS]M12
    End If

    If TempPlayer(index).tmpGuildInviteSlot > 0 Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M13")
        PlayerMsg Inviter_Index, Var1, BrightRed
        Exit Sub                    '[GREMIOS]M13
    End If

    'Permission 2 = Can Recruit
    If CheckGuildPermission(Inviter_Index, 2) = False Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M14")
        PlayerMsg Inviter_Index, Var1, BrightRed '[GREMIOS]M14
        Exit Sub
    End If
    
    TempPlayer(index).tmpGuildInviteSlot = GuildSlot
    '2 minute
    TempPlayer(index).tmpGuildInviteTimer = GetTickCount + 120000
    
    TempPlayer(index).tmpGuildInviteId = Player(Inviter_Index).GuildFileId
    
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M15")
    PlayerMsg Inviter_Index, Var1, Green '[GREMIOS]M15
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M18")
    PlayerMsg index, Trim$(Player(Inviter_Index).Name) & Var1 & GuildData(GuildSlot).Guild_Name & ".", Green '[GREMIOS]M18
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M16")
    PlayerMsg index, Var1, Green '[GREMIOS]M16
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M17")
    PlayerMsg index, Var1, Green '[GREMIOS]M17
End Sub

Public Sub Join_Guild(index As Long, GuildSlot As Long)
Dim OpenSlot As Long
    Dim Var1 As String
    
    If isPlaying(index) = False Then Exit Sub
    
    OpenSlot = FindOpenGuildMemberSlot(GuildSlot)
        'Guild full?
        If OpenSlot > 0 Then
        
            ' Check level
            If GetPlayerLevel(index) < Options.Join_Lvl Then    '[GREMIOS]M19
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M19")
                PlayerMsg index, Var1 & Options.Join_Lvl & "!", BrightRed
                Exit Sub
            End If
            
            ' Check if item is required
            If Not Options.Join_Item = 0 Then
                'Get item amount
                itemAmount = HasItem(index, Options.Join_Item)
                    
                ' Gold Req
                If itemAmount = 0 Or itemAmount < Options.Join_Cost Then    '[GREMIOS]M20
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M20")
                    PlayerMsg index, Var1 & Options.Join_Cost & " " & Item(Options.Join_Item).Name & "!", BrightRed
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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M9")
            PlayerMsg index, Var1 & GuildData(GuildSlot).Guild_Name & ".", BrightGreen
                            '[GREMIOS]M9
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M10")
            PlayerMsg index, Var1, BrightGreen
                            '[GREMIOS]M10
            'Update player to display guild name
            Call SendPlayerData(index)
            
        Else
            'Guild full display msg
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M21")
            PlayerMsg index, Var1, BrightRed '[GREMIOS]M21
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
    Dim i As Integer
    
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = False Then
            FindOpenGuildSlot = i
            Exit Function
        End If
        
        'No slot found how?
        FindOpenGuildSlot = 0
    Next i
End Function
Public Function FindOpenGuildMemberSlot(GuildSlot As Long) As Long
Dim i As Integer
    
    For i = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(i).Used = False Then
            FindOpenGuildMemberSlot = i
            Exit Function
        End If
    Next i
    
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
Dim i As Integer
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
    For i = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(i).Online = True Then
            GuildData(GuildSlot).Guild_Members(i).Online = False
        End If
    Next i
        
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
Dim i As Long

    For i = 1 To MAX_PLAYERS
        Call ClearGuild(i)
    Next i
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
Dim i As Integer
Dim UnloadGuild As Boolean

UnloadGuild = True

If GuildData(GuildSlot).In_Use = False Then Exit Sub

    For i = 1 To Player_HighIndex
        If isPlaying(i) Then
            If Player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                UnloadGuild = False
                Exit For
            End If
        End If
    Next i
    
    If UnloadGuild = True Then
        Call UnloadGuildSlot(GuildSlot)
    End If
End Sub
Public Sub GuildKick(GuildSlot As Long, index As Long, playerName As String)
Dim FoundOffline As Boolean
Dim IsOnline As Boolean
Dim OnlineIndex As Long
Dim MemberSlot As Long
Dim i As Integer
    Dim Var1 As String
    
    OnlineIndex = FindPlayer(playerName)
    
    If OnlineIndex = index Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M22")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M22
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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M23")
            PlayerMsg index, Var1, BrightRed '[GREMIOS]M23
            Exit Sub
        End If
        
        If GuildData(GuildSlot).Guild_Members(MemberSlot).Founder = True Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M24")
            PlayerMsg index, Var1, BrightRed '[GREMIOS]M24
            Exit Sub
        End If
        
        Player(OnlineIndex).GuildFileId = 0
        Player(OnlineIndex).GuildMemberId = 0
        TempPlayer(OnlineIndex).tmpGuildSlot = 0
        Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M25")
        PlayerMsg OnlineIndex, Var1, BrightRed '[GREMIOS]M25
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M26")
        PlayerMsg index, Var1, BrightRed                 '[GREMIOS]M26
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
        For i = 1 To MAX_GUILD_MEMBERS
            If playerName = Trim$(GuildData(GuildSlot).Guild_Members(i).User_Name) Then
                'Found them
                FoundOffline = True
                MemberSlot = i
                Exit For
            End If
        Next i
        
        If FoundOffline = True Then
        
            If MemberSlot = 0 Then Exit Sub
            
            Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
            Call SaveGuild(GuildSlot)
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M27")
            PlayerMsg index, Var1, BrightRed '[GREMIOS]M27
            Exit Sub
        End If
        
        If FoundOffline = False And IsOnline = False Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M28")
            PlayerMsg index, "" & playerName & Var1, BrightRed '[GREMIOS]M28
        End If
    
    End If
 
End Sub
Public Sub GuildLeave(index As Long)
Dim i As Integer
Dim GuildSlot As Long

    Dim Var1 As String
    
    'This is for the leave command only, kicking has its own sub because it handles both online and offline kicks, while this only handles online.
    
    If Not Player(index).GuildFileId > 0 Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M29")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M29
        Exit Sub
    End If
    
    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(Player(index).GuildMemberId).Founder = True Then
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M30")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M30
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M31")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M31
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
        
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M32")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M32

    Else
        'They don't match this slot remove them
        Player(index).GuildFileId = 0
        Player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
    End If
    
    
End Sub
Public Sub GuildLoginCheck(index As Long)
Dim i As Long
Dim GuildSlot As Long
Dim GuildLoaded As Boolean
GuildLoaded = False
Dim Var1 As String


    'Not in guild
    If Player(index).GuildFileId = 0 Then Exit Sub
    
    'Check to make sure the guild file exists
    If Not FileExist("\Data\guilds\Guild" & Player(index).GuildFileId & ".dat") Then
        'If guild was deleted remove user from guild
        Player(index).GuildFileId = 0
        Player(index).GuildMemberId = 0
        TempPlayer(index).tmpGuildSlot = 0
        Call SavePlayer(index)
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M33")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M33
        Exit Sub
    End If
    
    'First we need to see if our guild is loaded
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = True Then
            'If its already loaded set true
            If GuildData(i).Guild_Fileid = Player(index).GuildFileId Then
                GuildLoaded = True
                GuildSlot = i
                Exit For
            End If
        End If
    Next i
    
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
            PlayerMsg index, "" & GuildData(GuildSlot).Guild_MOTD, Blue
        End If
    End If
    
End Sub
Sub DisbandGuild(GuildSlot As Long, index As Long)
Dim i As Integer
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
            For i = 1 To Player_HighIndex
                If isPlaying(i) = True Then
                    If Player(i).GuildFileId = TmpGuildFileId Then
                        'remove from guild
                        Player(i).GuildFileId = 0
                        Player(i).GuildMemberId = 0
                        TempPlayer(i).tmpGuildSlot = 0
                        Call SavePlayer(i)
                        'Send player data so they don't have name over head anymore
                        Call SendPlayerData(i)
                    End If
                End If
            Next i
            
            'Unload Guild from memory
            Call UnloadGuildSlot(tmpGuildSlot)

            filename = App.Path & "\Data\Guilds\Guild" & TmpGuildFileId & ".dat"
            Kill filename
            
            
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M34")
            PlayerMsg index, Var1, BrightGreen '[GREMIOS]M34
         End If
    Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
        PlayerMsg index, Var1, BrightRed '[GREMIOS]M35
    End If
End Sub
Sub SendDataToGuild(ByVal GuildSlot As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If isPlaying(i) Then
            If Player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendGuild(ByVal SendToWholeGuild As Boolean, ByVal index As Long, ByVal GuildSlot)
    Dim buffer As clsBuffer
    Dim i As Integer
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
    For i = 1 To MAX_GUILD_MEMBERS
        buffer.WriteString GuildData(GuildSlot).Guild_Members(i).User_Name
        buffer.WriteInteger GuildData(GuildSlot).Guild_Members(i).Rank
        buffer.WriteString GuildData(GuildSlot).Guild_Members(i).Comment
        buffer.WriteByte GuildData(GuildSlot).Guild_Members(i).Online
    Next i
    
    'Send Ranks
    For i = 1 To MAX_GUILD_RANKS
            buffer.WriteString GuildData(GuildSlot).Guild_Ranks(i).Name
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            buffer.WriteByte GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b)
            buffer.WriteString Guild_Ranks_Premission_Names(b)
        Next b
    Next i
    
    If SendToWholeGuild = False Then
        SendDataTo index, buffer.ToArray()
    Else
        SendDataToGuild GuildSlot, buffer.ToArray()
    End If
    
    For i = 1 To MAX_GUILD_MEMBERS
        SendPlayerData i
    Next i
    
    Set buffer = Nothing
End Sub
Sub ToggleGuildAdmin(ByVal index As Long, ByVal GuildSlot, ByVal OpenAdmin As Boolean)
    Dim buffer As clsBuffer
    Dim i As Integer
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
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
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
Dim i As Integer

Dim Var1 As String

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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M36")
                    PlayerMsg index, Var1, BrightRed
                End If  '[GREMIOS]M36
            End If
        Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
            PlayerMsg index, Var1, BrightRed
        End If                 '[GREMIOS]M35
        HoldInt = 0
    Case 2
        'users
        If CheckGuildPermission(index, 5) = True Then
            'Guild Member Rank
            HoldInt = buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Rank = HoldInt
            Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M37")
                PlayerMsg index, Var1, BrightRed '[GREMIOS]M37
            End If
            
            'Guild Member Comment
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(SentIndex).Comment = buffer.ReadString
        Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
            PlayerMsg index, Var1, BrightRed
        End If                 '[GREMIOS]M35
        
    Case 3
        'ranks
        If CheckGuildPermission(index, 4) = True Then
            GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).Name = buffer.ReadString
                For i = 1 To MAX_GUILD_RANKS_PERMISSION
                    GuildData(TempPlayer(index).tmpGuildSlot).Guild_Ranks(SentIndex).RankPermission(i) = buffer.ReadByte
                Next i
        Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
            PlayerMsg index, Var1, BrightRed
        End If                 '[GREMIOS]M35
    
    End Select
    
    Call SendGuild(True, index, TempPlayer(index).tmpGuildSlot)
    
    Set buffer = Nothing
End Sub
Public Sub HandleGuildCommands(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Integer
    Dim SelectedIndex As Long
    Dim SendText As String
    Dim SendText2 As String
    Dim SelectedCommand As Integer
    Dim MembersCount As Long
    Dim buffer As clsBuffer
    Dim Var1 As String
    
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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M38")
            PlayerMsg index, Var1 & SendText & ".", BrightRed
        End If              '[GREMIOS]M38
        
    Case 3
        'leave
        Call GuildLeave(index)
        
    Case 4
        'admin
        If CheckGuildPermission(index, 1) = True Then
            Call ToggleGuildAdmin(index, TempPlayer(index).tmpGuildSlot, True)
        Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
            PlayerMsg index, Var1, BrightRed
        End If              '[GREMIOS]M35
    
    Case 5
        'view
        'This sets the default option
        If SendText = "" Then SendText = "en FLeer(FNum)"
        MembersCount = 0
        
        Select Case SendText
        Case "en FLeer(FNum)"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Online = True Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next i
            
            PlayerMsg index, "Total: " & MembersCount, Green
        
        Case "todos"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                    MembersCount = MembersCount + 1
                End If
            Next i
            
            PlayerMsg index, "Total: " & MembersCount, Green
        
        Case "fuera de FLeer(FNum)"
            PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).Online = False Then
                        PlayerMsg index, GuildData(TempPlayer(index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next i
            
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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M39")
                PlayerMsg index, Var1, BrightRed
            End If              '[GREMIOS]M39
        Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M40")
            PlayerMsg index, Var1, BrightRed
        End If              '[GREMIOS]M40
    Case 7
        'decline
        If TempPlayer(index).tmpGuildInviteSlot > 0 Then
            TempPlayer(index).tmpGuildInviteSlot = 0
            TempPlayer(index).tmpGuildInviteTimer = 0
            TempPlayer(index).tmpGuildInviteId = 0
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M41")
            PlayerMsg index, Var1, BrightRed
        Else                '[GREMIOS]M41
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M40")
            PlayerMsg index, Var1, BrightRed
        End If              '[GREMIOS]M40
        
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
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M42")
                    PlayerMsg index, Var1 & SendText & ".", BrightRed
                End If              '[GREMIOS]M42
            Else
                        Var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GREMIOS", "M35")
                 PlayerMsg index, Var1, BrightRed
            End If                  '[GREMIOS]M35
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
