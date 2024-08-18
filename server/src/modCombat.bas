Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > MAX_PLAYERS Then Exit Function
    Select Case Vital
        Case HP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 5 + 65
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Endurance) / 2)) * 15 + 150
            End Select
        Case MP
            Select Case GetPlayerClass(index)
                Case 1 ' Warrior
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
                Case 2 ' Mage
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 30 + 85
                Case Else ' Anything else - Warrior by default
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Intelligence) / 2)) * 5 + 25
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    ' Prevent subscript out of range
    If isPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (GetPlayerStat(index, Stats.Willpower) * 0.8) + 6
        Case MP
            I = (GetPlayerStat(index, Stats.Willpower) / 4) + 12.5
    End Select

    If I < 2 Then I = 2
    GetPlayerVitalRegen = I
End Function

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim weaponNum As Long
    
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If isPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    If GetPlayerEquipment(index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipment(index, Weapon)
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, strength) * Item(weaponNum).Data2 + (GetPlayerLevel(index) / 5)
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, strength) + (GetPlayerLevel(index) / 5)
    End If

End Function

Function GetNpcMaxVital(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim x As Long

    ' Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(npcNum).HP
        Case MP
            GetNpcMaxVital = 30 + (NPC(npcNum).stat(Intelligence) * 10) + 2
    End Select

End Function

Function GetNpcVitalRegen(ByVal npcNum As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    'Prevent subscript out of range
    If npcNum <= 0 Or npcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (NPC(npcNum).stat(Stats.Willpower) * 0.8) + 6
        Case MP
            I = (NPC(npcNum).stat(Stats.Willpower) / 4) + 12.5
    End Select
    
    GetNpcVitalRegen = I

End Function

Function GetNpcDamage(ByVal npcNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * NPC(npcNum).stat(Stats.strength) * NPC(npcNum).Damage + (NPC(npcNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanPlayerBlock = False

    rate = 0
    ' TODO : make it based on shield lulz
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanPlayerCrit = False

    rate = GetPlayerStat(index, Agility) / 52.08
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanPlayerCrit = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanPlayerDodge = False

    rate = GetPlayerStat(index, Agility) / 83.3
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanPlayerParry = False

    rate = GetPlayerStat(index, strength) * 0.25
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanPlayerParry = True
    End If
End Function

Public Function CanNpcBlock(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim stat As Long
Dim RndNum As Long

    CanNpcBlock = False
    
    stat = NPC(npcNum).stat(Stats.Agility) / 5  'guessed shield agility
    rate = stat / 12.08
    
    RndNum = rand(1, 100)
    
    If RndNum <= rate Then
        CanNpcBlock = True
    End If
    
End Function

Public Function CanNpcCrit(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanNpcCrit = False

    rate = NPC(npcNum).stat(Stats.Agility) / 52.08
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanNpcCrit = True
    End If
End Function

Public Function CanNpcDodge(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanNpcDodge = False

    rate = NPC(npcNum).stat(Stats.Agility) / 83.3
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcParry(ByVal npcNum As Long) As Boolean
Dim rate As Long
Dim RndNum As Long

    CanNpcParry = False

    rate = NPC(npcNum).stat(Stats.strength) * 0.25
    RndNum = rand(1, 100)
    If RndNum <= rate Then
        CanNpcParry = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal mapNpcNum As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long
Dim Dmg_Light As Long, Dmg_Dark As Long, Dmg_Neut As Long
Dim tempDmg As Long
Dim I As Long
Dim var1 As String

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, mapNpcNum) Then
    
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
        ' check if NPC can avoid the attack
        If CanNpcDodge(npcNum) Then
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A3")
            SendActionMsg mapnum, var1, Pink, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If
        If CanNpcParry(npcNum) Then
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A4")
            SendActionMsg mapnum, var1, Pink, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanNpcBlock(mapNpcNum)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (NPC(npcNum).stat(Stats.Agility) * 2))
        ' add combat weapon level if using weapon
        If GetPlayerEquipment(index, Weapon) > 0 Then Damage = Damage + rand(1, (GetPlayerCombatLevel(index, Item(GetPlayerEquipment(index, Weapon)).CombatTypeReq) * 2))
        
        'get damage boost from elements
        For I = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(index, I) > 0 Then
                With Item(GetPlayerEquipment(index, I))
                    If .Element_Light_Dmg > 0 Then
                        Dmg_Light = Dmg_Light + rand(1, .Element_Light_Dmg) 'randomize it
                    End If
                    If .Element_Dark_Dmg > 0 Then
                        Dmg_Dark = Dmg_Dark + rand(1, .Element_Dark_Dmg) 'randomize it
                    End If
                    If .Element_Neut_Dmg > 0 Then
                        Dmg_Neut = Dmg_Neut + rand(1, .Element_Neut_Dmg) 'randomize it
                    End If
                End With
            End If
        Next I
        
        Damage = Damage + Dmg_Light + Dmg_Dark + Dmg_Neut
        
        'get defense ability from npc
        With NPC(npcNum)
            If .Element_Light_Res > 0 Then
                If Dmg_Light > 0 Then
                    tempDmg = rand(1, .Element_Light_Res) 'randomize it
                    If tempDmg > Dmg_Light Then tempDmg = Dmg_Light
                    Damage = Damage - tempDmg
                End If
            End If
            If .Element_Dark_Res > 0 Then
                If Dmg_Dark > 0 Then
                    tempDmg = rand(1, .Element_Dark_Res) 'randomize it
                    If tempDmg > Dmg_Dark Then tempDmg = Dmg_Dark
                    Damage = Damage - tempDmg
                End If
            End If
            If .Element_Neut_Res > 0 Then
                If Dmg_Neut > 0 Then
                    tempDmg = rand(1, .Element_Neut_Res) 'randomize it
                    If tempDmg > Dmg_Neut Then tempDmg = Dmg_Neut
                    Damage = Damage - tempDmg
                End If
            End If
        End With
   
        'update damage
        Damage = Damage + tempDmg
        
        ' randomise from 1 to max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if it's a crit!
        If CanPlayerCrit(index) Then
            Damage = Damage * 1.5   '[COMBAT]A1
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A1")
            SendActionMsg mapnum, var1, BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
        End If
            
        If Damage > 0 Then
            Call PlayerAttackNpc(index, mapNpcNum, Damage)
        Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A2")
            Call PlayerMsg(index, var1, BrightRed) '[COMBAT]A2
        End If
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long
    Dim npcX As Long
    Dim npcY As Long
    Dim attackspeed As Long
    Dim GetWeapon As Byte
    Dim pX As Long, pY As Long

    ' Check for subscript out of range
    If isPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(attacker)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_FRIENDLY Then
            If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                Exit Function
            End If
        End If
    End If

    ' Make sure they are on the same map
    If isPlaying(attacker) Then
    
        ' exit out early
        If IsSpell Then
             If npcNum > 0 Then
                If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                    CanPlayerAttackNpc = True
                    Exit Function
                End If
            End If
        End If

        ' attack speed from weapon
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            attackspeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
        Else
            attackspeed = 1000
        End If

        If npcNum > 0 And GetTickCount > TempPlayer(attacker).AttackTimer + attackspeed Then
            ' Check if at same coordinates
            Select Case GetPlayerDir(attacker)
                Case DIR_UP
                    npcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    npcY = MapNpc(mapnum).NPC(mapNpcNum).y + 1
                Case DIR_DOWN
                    npcX = MapNpc(mapnum).NPC(mapNpcNum).x
                    npcY = MapNpc(mapnum).NPC(mapNpcNum).y - 1
                Case DIR_LEFT
                    npcX = MapNpc(mapnum).NPC(mapNpcNum).x + 1
                    npcY = MapNpc(mapnum).NPC(mapNpcNum).y
                Case DIR_RIGHT
                    npcX = MapNpc(mapnum).NPC(mapNpcNum).x - 1
                    npcY = MapNpc(mapnum).NPC(mapNpcNum).y
            End Select
            
            GetWeapon = GetPlayerEquipment(attacker, Weapon)
            If GetWeapon = 0 Then GetWeapon = 1
                If Not Item(GetWeapon).CombatTypeReq = 5 Then ' Check for Polearm
                    If npcX = GetPlayerX(attacker) Then
                        If npcY = GetPlayerY(attacker) Then
                            If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                                CanPlayerAttackNpc = True
                            ElseIf NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                                Call CheckTasks(attacker, QUEST_TYPE_GOTALK, npcNum)
                                Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, npcNum)
                                Call CheckTasks(attacker, QUEST_TYPE_GOGET, npcNum)
                                
                                If NPC(npcNum).Quest = YES Then
                                    If Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED Then
                                        If Quest(NPC(npcNum).Quest).Repeat = YES Then
                                            Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED_BUT
                                            Exit Function
                                        End If
                                    End If
                                    
                                    If CanStartQuest(attacker, NPC(npcNum).QuestNum) Then
                                        'if can start show the request message (speech1)
                                        QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(1)), NPC(npcNum).QuestNum
                                        Exit Function
                                    End If
                                    
                                    If QuestInProgress(attacker, NPC(npcNum).QuestNum) Then
                                        'if the quest is in progress show the meanwhile message (speech2)
                                        QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(2)), 0
                                        Exit Function
                                    End If
                                End If
                                
                                If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                    Call SendChatBubble(mapnum, mapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                                End If
                            End If
                        End If
                    End If
                Else ' Item is Polearm so range is 2 tiles
                    If ((npcX = GetPlayerX(attacker) Or npcX + 1 = GetPlayerX(attacker) Or npcX - 1 = GetPlayerX(attacker)) And npcY = GetPlayerY(attacker)) Or ((npcY = GetPlayerY(attacker) Or npcY + 1 = GetPlayerY(attacker) Or npcY - 1 = GetPlayerY(attacker)) And npcX = GetPlayerX(attacker)) Then
                        If NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(npcNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
                            ' Make sure tile we're trying to swing through isn't a blocked tile or a resource
                            Select Case GetPlayerDir(attacker)
                                Case DIR_UP
                                    pX = 0
                                    pY = -1
                                Case DIR_DOWN
                                    pX = 0
                                    pY = 1
                                Case DIR_LEFT
                                    pX = -1
                                    pY = 0
                                Case DIR_RIGHT
                                    pX = 1
                                    pY = 0
                            End Select
                            
                            If Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker) + pX, GetPlayerY(attacker) + pY).Type <> TILE_TYPE_BLOCKED And Map(GetPlayerMap(attacker)).Tile(GetPlayerX(attacker) + pX, GetPlayerY(attacker) + pY).Type <> TILE_TYPE_RESOURCE Then
                                CanPlayerAttackNpc = True
                            End If
                        ElseIf NPC(npcNum).Behaviour = NPC_BEHAVIOUR_FRIENDLY Or NPC(npcNum).Behaviour = NPC_BEHAVIOUR_SHOPKEEPER Then
                            Call CheckTasks(attacker, QUEST_TYPE_GOTALK, npcNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGIVE, npcNum)
                            Call CheckTasks(attacker, QUEST_TYPE_GOGET, npcNum)
                                
                            If NPC(npcNum).Quest = YES Then
                                If Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED Then
                                    If Quest(NPC(npcNum).Quest).Repeat = YES Then
                                        Player(attacker).PlayerQuest(NPC(npcNum).Quest).Status = QUEST_COMPLETED_BUT
                                        Exit Function
                                    End If
                                End If
                                    
                                If CanStartQuest(attacker, NPC(npcNum).QuestNum) Then
                                    'if can start show the request message (speech1)
                                    QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(1)), NPC(npcNum).QuestNum
                                    Exit Function
                                End If
                                    
                                If QuestInProgress(attacker, NPC(npcNum).QuestNum) Then
                                    'if the quest is in progress show the meanwhile message (speech2)
                                    QuestMessage attacker, NPC(npcNum).QuestNum, Trim$(Quest(NPC(npcNum).QuestNum).Speech(2)), 0
                                    Exit Function
                                End If
                            End If
                                
                            If Len(Trim$(NPC(npcNum).AttackSay)) > 0 Then
                                Call SendChatBubble(mapnum, mapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(npcNum).AttackSay), DarkBrown)
                            End If
                        End If
                    End If
                End If
            'End If
            
        End If
    End If

End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal mapNpcNum As Long, ByVal Damage As Long, Optional ByVal spellnum As Long, Optional ByVal overTime As Boolean = False)
    Dim Name As String
    Dim EXP As Long
    Dim n As Long
    Dim I As Long
    Dim STR As Long
    Dim DEF As Long
    Dim mapnum As Long
    Dim npcNum As Long
    Dim num As Long
    Dim buffer As clsBuffer
    Dim DropVal As Long
    Dim DropMulti As Double

    ' Check for subscript out of range
    If isPlaying(attacker) = False Or mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(attacker)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    If npcNum < 1 Then Exit Sub
    Name = Trim$(NPC(npcNum).Name)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    ' Check for a weapon and say damage
        If Damage <= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
            SendActionMsg mapnum, "-" & Damage, BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        Else
            SendActionMsg mapnum, "-" & MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If
        SendBlood GetPlayerMap(attacker), MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y
    
    ' send animation
    If n > 0 Then
        If Not overTime Then
            If spellnum = 0 Then
                Call SendAnimation(mapnum, Item(n).Animation, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                SendMapSound attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seItem, n
            End If
        End If
    End If
    If spellnum > 0 Then
        Call SendAnimation(mapnum, Spell(spellnum).SpellAnim, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
        SendMapSound attacker, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, SoundEntity.seSpell, spellnum
    End If
    
    Call SendFlash(mapNpcNum, mapnum, True)
    
    If Damage >= MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) Then
        
        ' Calculate exp to give attacker
        If NPC(npcNum).RandExp = 0 Then
            EXP = NPC(npcNum).EXP
        Else
            'randomize exp within specified value
            If NPC(npcNum).Percent_5 = 1 Then
                EXP = rand(NPC(npcNum).EXP - (NPC(npcNum).EXP * 0.05), NPC(npcNum).EXP + (NPC(npcNum).EXP * 0.05))
            ElseIf NPC(npcNum).Percent_10 = 1 Then
                EXP = rand(NPC(npcNum).EXP - (NPC(npcNum).EXP * 0.1), NPC(npcNum).EXP + (NPC(npcNum).EXP * 0.1))
            ElseIf NPC(npcNum).Percent_20 = 1 Then
                EXP = rand(NPC(npcNum).EXP - (NPC(npcNum).EXP * 0.2), NPC(npcNum).EXP + (NPC(npcNum).EXP * 0.2))
            End If
        End If
        
        ' Double exp?
        If DoubleExp Then EXP = EXP * 2

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If
        
        ' register kill
        Call SetPlayerKills(attacker, 1)
        
        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, EXP, attacker, GetPlayerMap(attacker)
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, EXP
        End If
        
        ' Give player combat exp
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If Item(GetPlayerEquipment(attacker, Weapon)).CombatTypeReq > 0 Then
                GivePlayerCombatEXP attacker, GetPlayerEquipment(attacker, Weapon), EXP / 4
            End If
        End If
                
        ' Check if the player is in a party!
        If TempPlayer(attacker).inParty > 0 Then
            num = rand(1, Party(TempPlayer(attacker).inParty).MemberCount)
            Do While GetPlayerMap(Party(TempPlayer(attacker).inParty).Member(num)) <> mapnum
                num = rand(1, Party(TempPlayer(attacker).inParty).MemberCount) ' Randomly pick party member on same map
            Loop
            'Drop the goods if they get it
            For I = 1 To MAX_NPC_DROP_ITEMS
                n = Int(Rnd * NPC(npcNum).Drops(I).DropChance) + 1
                If n = 1 Then
                    Call GiveInvItem(Party(TempPlayer(attacker).inParty).Member(num), NPC(npcNum).Drops(I).DropItem, NPC(npcNum).Drops(I).DropItemValue, True)
                  '  Call PartyMsg(TempPlayer(attacker).inParty, GetPlayerName(Party(TempPlayer(attacker).inParty).Member(num)) & " looted " & Trim$(Item(NPC(npcNum).Drops(I).DropItem).Name) & ".", Yellow)
                  ' Parche Fix 0.4 Party Drop Objetos
                End If
            Next I
        Else
            For I = 1 To MAX_NPC_DROP_ITEMS
                'Drop the goods if they get it
                n = Int(Rnd * NPC(npcNum).Drops(I).DropChance) + 1
                If n = 1 Then
                    DropVal = NPC(npcNum).Drops(I).DropItemValue
                    If NPC(npcNum).Drops(I).RandCurrency Then
                        If NPC(npcNum).Drops(I).P_5 Then DropMulti = 0.05
                        If NPC(npcNum).Drops(I).P_10 Then DropMulti = 0.1
                        If NPC(npcNum).Drops(I).P_20 Then DropMulti = 0.2
                        If DropMulti > 0 Then DropVal = rand(DropVal - (DropVal * DropMulti), DropVal + (DropVal * DropMulti))
                    End If
                    Call SpawnItem(NPC(npcNum).Drops(I).DropItem, DropVal, mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y)
                End If
            Next I
        End If

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(mapnum).NPC(mapNpcNum).num = 0
        MapNpc(mapnum).NPC(mapNpcNum).SpawnWait = GetTickCount
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = 0
        UpdateMapBlock mapnum, MapNpc(mapnum).NPC(mapNpcNum).x, MapNpc(mapnum).NPC(mapNpcNum).y, False
        
        ' clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNpc(mapnum).NPC(mapNpcNum).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(mapnum).NPC(mapNpcNum).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        ' ONNPCDEATH NPC DEATH
        
       
        Call CheckTasks(attacker, QUEST_TYPE_GOSLAY, npcNum)
        ' send death to the map
        Set buffer = New clsBuffer
        buffer.WriteLong SNpcDead
        buffer.WriteLong mapNpcNum
        SendDataToMap mapnum, buffer.ToArray()
        Set buffer = Nothing
        
        'Loop through entire map and purge NPC from targets
        For I = 1 To Player_HighIndex
            If isPlaying(I) And IsConnected(I) Then
                If Player(I).Map = mapnum Then
                    If TempPlayer(I).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(I).target = mapNpcNum Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) = MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) - Damage

        ' Set the NPC target to the player
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 1 ' player
        MapNpc(mapnum).NPC(mapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNpc(mapnum).NPC(I).num = MapNpc(mapnum).NPC(mapNpcNum).num Then
                    MapNpc(mapnum).NPC(I).target = attacker
                    MapNpc(mapnum).NPC(I).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
        MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning spell, stun the npc
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunNPC mapNpcNum, mapnum, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Npc mapnum, mapNpcNum, spellnum, attacker
            End If
        End If
        
        SendMapNpcVitals mapnum, mapNpcNum
    End If

    If spellnum = 0 Then
        ' Reset attack timer
        TempPlayer(attacker).AttackTimer = GetTickCount
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long)
Dim mapnum As Long, npcNum As Long, blockAmount As Long, Damage As Long
Dim I As Long
Dim Dmg_Light As Long, Dmg_Dark As Long, Dmg_Neut As Long
Dim tempDmg As Long

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(mapNpcNum, index) Then
        mapnum = GetPlayerMap(index)
        npcNum = MapNpc(mapnum).NPC(mapNpcNum).num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg mapnum, "Esquivo!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If
        If CanPlayerParry(index) Then
            SendActionMsg mapnum, "Desviado!", Pink, 1, (Player(index).x * 32), (Player(index).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(npcNum)
        
        ' if the player blocks, take away the block amount
        blockAmount = CanPlayerBlock(index)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(index, Agility) * 2))
        
        'get damage boost
        With NPC(npcNum)
            If .Element_Light_Dmg > 0 Then
                Dmg_Light = Dmg_Light + rand(1, .Element_Light_Dmg) 'randomize it
            End If
            If .Element_Dark_Dmg > 0 Then
                Dmg_Dark = Dmg_Dark + rand(1, .Element_Dark_Dmg) 'randomize it
            End If
            If .Element_Neut_Dmg > 0 Then
                Dmg_Neut = Dmg_Neut + rand(1, .Element_Neut_Dmg) 'randomize it
            End If
        End With
        
        Damage = Damage + Dmg_Light + Dmg_Dark + Dmg_Neut
        
        'get damage resist from player's elements
        For I = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(index, I) > 0 Then
                With Item(GetPlayerEquipment(index, I))
                    'recalculate damage -'s
                    If .Element_Light_Dmg > 0 Then
                        If NPC(npcNum).Element_Light_Dmg > 0 Then
                            tempDmg = rand(1, .Element_Light_Res) 'randomize it
                            If tempDmg > Dmg_Light Then tempDmg = Dmg_Light
                            Damage = Damage - tempDmg
                        End If
                    End If
                    If .Element_Dark_Dmg > 0 Then
                        If NPC(npcNum).Element_Dark_Dmg > 0 Then
                            Damage = Damage - rand(1, .Element_Dark_Res) 'randomize it
                            If tempDmg > Dmg_Dark Then tempDmg = Dmg_Dark
                            Damage = Damage - tempDmg
                        End If
                    End If
                    If .Element_Neut_Dmg > 0 Then
                        If NPC(npcNum).Element_Neut_Dmg > 0 Then
                            Damage = Damage - rand(1, .Element_Neut_Res) 'randomize it
                            If tempDmg > Dmg_Neut Then tempDmg = Dmg_Neut
                            Damage = Damage - tempDmg
                        End If
                    End If
                End With
            End If
        Next I
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if crit hit
        If CanNpcCrit(npcNum) Then
            Damage = Damage * 1.5 '[COMBAT]A1
            Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A1")
            SendActionMsg mapnum, var1, BrightCyan, 1, (MapNpc(mapnum).NPC(mapNpcNum).x * 32), (MapNpc(mapnum).NPC(mapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(mapNpcNum, index, Damage)
        End If
    End If
End Sub

Function CanNpcAttackPlayer(ByVal mapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum As Long
    Dim npcNum As Long

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or Not isPlaying(index) Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(mapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
    npcNum = MapNpc(mapnum).NPC(mapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNpc(mapnum).NPC(mapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(mapnum).NPC(mapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    MapNpc(mapnum).NPC(mapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If isPlaying(index) Then
        If npcNum > 0 Then

            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(index) = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNpc(mapnum).NPC(mapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(mapnum).NPC(mapNpcNum).x) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal mapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim EXP As Long
    Dim mapnum As Long
    Dim I As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If mapNpcNum <= 0 Or mapNpcNum > MAX_MAP_NPCS Or isPlaying(victim) = False Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).num <= 0 Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(victim)
    Name = Trim$(NPC(MapNpc(mapnum).NPC(mapNpcNum).num).Name)
    
    ' Send this packet so they can see the npc attacking
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcAttack
    buffer.WriteLong mapNpcNum
    SendDataToMap mapnum, buffer.ToArray()
    Set buffer = Nothing
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    MapNpc(mapnum).NPC(mapNpcNum).stopRegen = True
    MapNpc(mapnum).NPC(mapNpcNum).stopRegenTimer = GetTickCount
    
    ' Say damage
    ' New way to visualize health decreasing. Dynamic color
    ' Green for 70%+ health
    
    If GetPlayerVital(victim, Vitals.HP) - Damage > GetPlayerMaxVital(victim, Vitals.HP) * 0.7 Then
        SendActionMsg GetPlayerMap(victim), GetPlayerVital(victim, Vitals.HP) & " -" & Damage, Green, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    ' Yellow for 35%+ health
    ElseIf GetPlayerVital(victim, Vitals.HP) - Damage > GetPlayerMaxVital(victim, Vitals.HP) * 0.35 Then
        SendActionMsg GetPlayerMap(victim), GetPlayerVital(victim, Vitals.HP) & " -" & Damage, Yellow, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    Else
    ' Red for anything lower
        SendActionMsg GetPlayerMap(victim), GetPlayerVital(victim, Vitals.HP) & " -" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    End If
    
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    
    Call SendAnimation(mapnum, NPC(MapNpc(GetPlayerMap(victim)).NPC(mapNpcNum).num).Animation, GetPlayerX(victim), GetPlayerY(victim), TARGET_TYPE_PLAYER, victim)
    ' send the sound
    SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(mapnum).NPC(mapNpcNum).num
    
    Call SendFlash(victim, mapnum, False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " fue asesinado por " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(mapnum).NPC(mapNpcNum).target = 0
        MapNpc(mapnum).NPC(mapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
Dim blockAmount As Long
Dim npcNum As Long
Dim mapnum As Long
Dim Damage As Long, tempDmg As Long
Dim Dmg_Light As Long, Dmg_Dark As Long, Dmg_Neut As Long
Dim I As Long
Dim var1 As String

    Damage = 0

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
    
        mapnum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then      '[COMBAT]A3
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A3")
            SendActionMsg mapnum, var1, Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If
        If CanPlayerParry(victim) Then      '[COMBAT]A4
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A4")
            SendActionMsg mapnum, var1, Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        blockAmount = CanPlayerBlock(victim)
        Damage = Damage - blockAmount
        
        ' take away armour
        Damage = Damage - rand(1, (GetPlayerStat(victim, Agility) * 2))
        
        'get damage from player elements and defence from opponents elements
        For I = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(attacker, I) > 0 Then
                With Item(GetPlayerEquipment(attacker, I))
                
                    ' Light damage and opponents defence to it
                    If .Element_Light_Dmg > 0 Then
                        Dmg_Light = rand(1, .Element_Light_Dmg)
                        Damage = Damage + Dmg_Light  'randomize it
                        
                        If GetPlayerEquipment(victim, I) > 0 Then
                            If Item(GetPlayerEquipment(victim, I)).Element_Light_Res > 0 Then
                                tempDmg = rand(1, .Element_Light_Res) 'randomize it
                                If tempDmg > Dmg_Light Then tempDmg = Dmg_Light
                                Damage = Damage - tempDmg
                            End If
                        End If
                    End If
                    
                    ' Dark damage and opponents defence to it
                    If .Element_Dark_Dmg > 0 Then
                        Dmg_Dark = rand(1, .Element_Dark_Dmg)
                        Damage = Damage + Dmg_Dark  'randomize it
                        
                        If GetPlayerEquipment(victim, I) > 0 Then
                            If Item(GetPlayerEquipment(victim, I)).Element_Dark_Res > 0 Then
                                tempDmg = rand(1, .Element_Dark_Res) 'randomize it
                                If tempDmg > Dmg_Dark Then tempDmg = Dmg_Dark
                                Damage = Damage - tempDmg
                            End If
                        End If
                    End If
                    
                    ' Neutral damage and opponents defence to it
                    If .Element_Neut_Dmg > 0 Then
                        Dmg_Neut = rand(1, .Element_Neut_Dmg)
                        Damage = Damage + Dmg_Dark  'randomize it
                        
                        If GetPlayerEquipment(victim, I) > 0 Then
                            If Item(GetPlayerEquipment(victim, I)).Element_Neut_Res > 0 Then
                                tempDmg = rand(1, .Element_Neut_Res) 'randomize it
                                If tempDmg > Dmg_Neut Then tempDmg = Dmg_Neut
                                Damage = Damage - tempDmg
                            End If
                        End If
                    End If
                End With
            End If
        Next I
        
        ' randomise for up to 10% lower than max hit
        Damage = rand(1, Damage)
        
        ' * 1.5 if can crit
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5       '[COMBAT]A1
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A1")
            SendActionMsg mapnum, var1, BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else                            '[COMBAT]A2
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "A2")
            Call PlayerMsg(attacker, var1, BrightRed)
        End If
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal IsProjectile As Boolean = False) As Boolean
Dim var1 As String

    If Not IsSpell And Not IsProjectile Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Check for subscript out of range
    If Not isPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell And Not IsProjectile Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP
    
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN
    
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
    
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M1")
            Call PlayerMsg(attacker, var1, BrightRed) '[COMBAT]M1
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M2")
        Call PlayerMsg(attacker, var1, BrightBlue) '[COMBAT]M2
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M3")
        Call PlayerMsg(attacker, var1 & GetPlayerName(victim) & ".", BrightRed) '[COMBAT]M3
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M4")
        Call PlayerMsg(attacker, var1, BrightRed) '[COMBAT]M4
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M5")
        Call PlayerMsg(attacker, GetPlayerName(victim) & var1, BrightRed) '[COMBAT]M5
        Exit Function
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal spellnum As Long = 0)
    Dim EXP As Long
    Dim n As Long
    Dim I As Long
    Dim buffer As clsBuffer
       Dim var1 As String
       
    ' Check for subscript out of range
    If isPlaying(attacker) = False Or isPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
    
       If Damage <> 0 Then
    SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
    SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
    End If
    ' send animation
    If n > 0 Then
        If spellnum = 0 Then
            Call SendAnimation(GetPlayerMap(victim), Item(n).Animation, GetPlayerX(victim), GetPlayerY(victim))
            SendMapSound attacker, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seItem, n
        End If
    End If
    
    If spellnum > 0 Then
        Call SendAnimation(GetPlayerMap(victim), Spell(spellnum).SpellAnim, GetPlayerX(victim), GetPlayerY(victim))
        SendMapSound attacker, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, spellnum
    End If
    
    Call SendFlash(victim, GetPlayerMap(victim), False)
    
    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Player is dead
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M6")
        Call GlobalMsg(GetPlayerName(victim) & var1 & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker            '[COMBAT]M6
        EXP = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If

        If EXP = 0 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M7")
            Call PlayerMsg(victim, var1, BrightRed)    '[COMBAT]M7
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M8")
            Call PlayerMsg(attacker, var1, BrightBlue) '[COMBAT]M8
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - EXP)
            SendEXP victim
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M9")
            Call PlayerMsg(victim, var1 & EXP & " exp.", BrightRed) '[COMBAT]M9
            
            ' check if we're in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, EXP, attacker, GetPlayerMap(attacker)
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, EXP
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For I = 1 To Player_HighIndex
            If isPlaying(I) And IsConnected(I) Then
                If Player(I).Map = GetPlayerMap(attacker) Then
                    If TempPlayer(I).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(I).target = victim Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendTarget I
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then 'EaSee 0.6 Facciones
            If GetPlayerPK(attacker) = NO Then
            If Class(GetPlayerClass(victim)).Faccion(0) = Class(GetPlayerClass(attacker)).Faccion(0) Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)                   '[COMBAT]M10
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M10")
                Call GlobalMsg(GetPlayerName(attacker) & var1, BrightRed)
            Else
            
            Call GiveInvItem(attacker, Class(GetPlayerClass(attacker)).ItemFaccion, 1, True)
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M11")
            Call PlayerMsg(attacker, var1 & Item(Class(GetPlayerClass(attacker)).ItemFaccion).Name, BrightRed)
                                        '[COMBAT]M11
            End If
            End If

        Else                                            '[COMBAT]M12
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M12")
            Call GlobalMsg(GetPlayerName(victim) & var1, BrightRed)
        End If
        
        Call CheckTasks(attacker, QUEST_TYPE_GOKILL, victim)
        Call OnDeath(victim)
    Else
        ' Player not dead, just do the damage
        If spellnum <> 0 Then
        If Spell(spellnum).Type <> SPELL_TYPE_BUFF Then
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        End If
        
        If Spell(spellnum).Type = SPELL_TYPE_BUFF Then
            'Inversion - Confusion
            
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Inversion > 0 Then Confusion victim, spellnum 'Confusion
            'Paralisis
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Paralisis > 0 Then Paralizar victim, spellnum
            'Envenenar
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Veneno > 0 Then Envenenar victim, spellnum
            'Velocidad
            If Spell(spellnum).Duration > 0 And Spell(spellnum).VelocidadCaminar2 > 0 Or Spell(spellnum).VelocidadCorrer2 > 0 Then Velocidad victim, spellnum
            'Fuerza
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Fuerza > 0 Then
            Fuerza victim, spellnum
            End If
            'Destreza
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Destreza > 0 Then Destreza victim, spellnum
            'Agilidad
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Agilidad > 0 Then Agilidad victim, spellnum
            'Inteligencia
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Inteligencia > 0 Then Inteligencia victim, spellnum
            'Voluntad
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Voluntad > 0 Then Voluntad victim, spellnum
            'Sprite
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Sprite > 0 Then Sprite victim, spellnum
            'Invisibilidad
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Invisibilidad > 0 Then
            Invisibilidad victim, spellnum
            End If
            'Transportar
            If Spell(spellnum).Transportar > 0 And Spell(spellnum).TransportarMapa > 0 And Spell(spellnum).TransportarX > 0 And Spell(spellnum).TransportarY > 0 Then TransportarH victim, spellnum
        End If
        End If
        
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        'if a stunning spell, stun the player
        If spellnum > 0 Then
            If Spell(spellnum).StunDuration > 0 Then StunPlayer victim, spellnum
            ' DoT
            If Spell(spellnum).Duration > 0 Then
                AddDoT_Player victim, spellnum, attacker
            End If
        End If
        
        
        'NO
        End If
        'NO
    

    ' Reset attack timer
    TempPlayer(attacker).AttackTimer = GetTickCount
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal spellslot As Long)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    Dim var1 As String
    
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)
    
    If spellnum <= 0 Or spellnum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(spellslot) > GetTickCount Then
        'Hechizos recargando!
        PlayerMsg index, GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M13"), BrightRed
        Exit Sub            '[COMBAT]M13
    End If

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M14")
        Call PlayerMsg(index, var1, BrightRed)
        Exit Sub                '[COMBAT]M14
    End If
    
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M15")
        Call PlayerMsg(index, var1 & LevelReq & ".", BrightRed)
        Exit Sub                '[COMBAT]M15
    End If
    
    AccessReq = Spell(spellnum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M16")
        Call PlayerMsg(index, var1, BrightRed)
        Exit Sub                    '[COMBAT]M16
    End If
    
    ClassReq = Spell(spellnum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M17")
            Call PlayerMsg(index, var1 & CheckGrammar(Trim$(Class(ClassReq).Name)) & ".", BrightRed)
            Exit Sub                '[COMBAT]M17
        End If
    End If
    
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            'ESTE TOCA
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            'SEGUNDO INTENTO ESTE
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    Range = Spell(spellnum).Range
    HasBuffered = False
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    
    'change the target if we're casting an AOE and have no target
    If SpellCastType = 3 Then
        If targetType = 0 Then targetType = TARGET_TYPE_NPC
        If target = 0 Then target = FindNPCInRange(index, mapnum, Range)
        'if it's still 0
        If target = 0 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M18")
            PlayerMsg index, var1, BrightRed    '[COMBAT]M18
            SendClearSpellBuffer index
            Exit Sub
        End If
    End If
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                'No tienes objetivo.
                PlayerMsg index, GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M19"), BrightRed    '[COMBAT]M19
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    'Objetivo fuera de rango.
                    PlayerMsg index, GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M20"), BrightRed '[COMBAT]M20
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP And Spell(spellnum).Type <> SPELL_TYPE_BUFF Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(mapnum).NPC(target).x, MapNpc(mapnum).NPC(target).y) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M20")
                    PlayerMsg index, var1, BrightRed '[COMBAT]M20
                    HasBuffered = False
                Else
                    ' go through spell types
                    If Spell(spellnum).Type <> SPELL_TYPE_DAMAGEHP And Spell(spellnum).Type <> SPELL_TYPE_DAMAGEMP And Spell(spellnum).Type <> SPELL_TYPE_BUFF Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then         '[COMBAT]M21
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M21")
        SendAnimation mapnum, Spell(spellnum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg mapnum, var1 & Trim$(Spell(spellnum).Name) & ".", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).spellBuffer.Spell = spellslot
        TempPlayer(index).spellBuffer.Timer = GetTickCount
        TempPlayer(index).spellBuffer.target = TempPlayer(index).target
        TempPlayer(index).spellBuffer.tType = TempPlayer(index).targetType
        
        
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim spellnum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim mapnum As Long
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim I As Long
    Dim AoE As Long
    Dim Range As Long
    Dim VitalType As Byte
    Dim increment As Boolean
    Dim x As Long, y As Long
   Dim var1 As String
   
    Dim buffer As clsBuffer
    Dim SpellCastType As Long
   
    DidCast = False

    ' Prevent subscript out of range
    If spellslot <= 0 Or spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    spellnum = GetPlayerSpell(index, spellslot)
    mapnum = GetPlayerMap(index)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then Exit Sub

    MPCost = Spell(spellnum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M14")
        Call PlayerMsg(index, var1, BrightRed) '[COMBAT]M14
        Exit Sub
    End If
   
    LevelReq = Spell(spellnum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M15")
        Call PlayerMsg(index, var1 & LevelReq & ".", BrightRed)
        Exit Sub                '[COMBAT]M15
    End If
   
    AccessReq = Spell(spellnum).AccessReq
   
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M16")
        Call PlayerMsg(index, var1, BrightRed)
        Exit Sub                '[COMBAT]M16
    End If
   
    ClassReq = Spell(spellnum).ClassReq
   
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M17")
            Call PlayerMsg(index, var1 & CheckGrammar(Trim$(Class(ClassReq).Name)) & ".", BrightRed)
            Exit Sub            '[COMBAT]M17
        End If
    End If
   
    ' find out what kind of spell it is! self cast, target or AOE
    If Spell(spellnum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(spellnum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
   
    ' set the vital
    Vital = Spell(spellnum).Vital
    
    ' set elemental dmg
    Vital = Vital + rand(1, Spell(spellnum).Dmg_Light)
    Vital = Vital + rand(1, Spell(spellnum).Dmg_Dark)
    Vital = Vital + rand(1, Spell(spellnum).Dmg_Neut)
    
    
    AoE = Spell(spellnum).AoE
    Range = Spell(spellnum).Range
    
    'change the target if we're casting an AOE and have no target
    If SpellCastType = 3 Then
        If targetType = 0 Then targetType = TARGET_TYPE_NPC
        If target = 0 Then target = FindNPCInRange(index, mapnum, Range)
        'if it's still 0
        If target = 0 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M18")
            PlayerMsg index, var1, BrightRed
            Exit Sub            '[COMBAT]M18
        End If
    End If
    
    If target > 0 Then
        If targetType = TARGET_TYPE_PLAYER Then
            For I = 1 To Equipment.Equipment_Count - 1
                If GetPlayerEquipment(target, I) > 0 Then
                    With Item(GetPlayerEquipment(target, I))
                        If .Element_Light_Res > 0 Then
                            If Spell(spellnum).Dmg_Light > 0 Then
                                Vital = Vital - rand(1, .Element_Light_Res)
                            End If
                        End If
                        
                        If .Element_Dark_Res > 0 Then
                            If Spell(spellnum).Dmg_Dark > 0 Then
                                Vital = Vital - rand(1, .Element_Dark_Res)
                            End If
                        End If
                        
                        If .Element_Neut_Res > 0 Then
                            If Spell(spellnum).Dmg_Neut > 0 Then
                                Vital = Vital - rand(1, .Element_Neut_Res)
                            End If
                        End If
                    End With
                End If
            Next I
        ElseIf targetType = TARGET_TYPE_NPC Then
            With NPC(target)
                If .Element_Light_Res > 0 Then
                    If Spell(spellnum).Dmg_Light > 0 Then
                        Vital = Vital - rand(1, .Element_Light_Res)
                    End If
                End If
                
                If .Element_Dark_Res > 0 Then
                    If Spell(spellnum).Dmg_Dark > 0 Then
                        Vital = Vital - rand(1, .Element_Dark_Res)
                    End If
                End If
                
                If .Element_Neut_Res > 0 Then
                    If Spell(spellnum).Dmg_Neut > 0 Then
                        Vital = Vital - rand(1, .Element_Neut_Res)
                    End If
                End If
            End With
        End If
    End If
   
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, spellnum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(spellnum).Map, Spell(spellnum).x, Spell(spellnum).y
                    SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    DidCast = True
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
               
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(mapnum).NPC(target).x
                    y = MapNpc(mapnum).NPC(target).y
                End If
               
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M19")
                    PlayerMsg index, var1, BrightRed
                    SendClearSpellBuffer index '[COMBAT]M19
                End If
            End If
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If isPlaying(I) Then
                            If I <> index Then
                                If GetPlayerMap(I) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(index, I, True) Then
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            PlayerAttackPlayer index, I, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(I).num > 0 Then
                            If MapNpc(mapnum).NPC(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(I).x, MapNpc(mapnum).NPC(I).y) Then
                                    If CanPlayerAttackNpc(index, I, True) Then
                                        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                        PlayerAttackNpc index, I, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                   
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If isPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(index) Then
                                If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                    SpellPlayer_Effect VitalType, increment, I, Vital, spellnum
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(I).num > 0 Then
                            If MapNpc(mapnum).NPC(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(I).x, MapNpc(mapnum).NPC(I).y) Then
                                    SpellNpc_Effect VitalType, increment, I, Vital, spellnum, mapnum
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_BUFF 'EaSee 0.5
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If isPlaying(I) Then
                            If I <> index Then
                                If GetPlayerMap(I) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(index, I, True) Then
                                            SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            PlayerAttackPlayer index, I, Vital, spellnum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(mapnum).NPC(I).num > 0 Then
                            If MapNpc(mapnum).NPC(I).Vital(HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(mapnum).NPC(I).x, MapNpc(mapnum).NPC(I).y) Then
                                    If CanPlayerAttackNpc(index, I, True) Then
                                        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                        PlayerAttackNpc index, I, Vital, spellnum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
           
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(mapnum).NPC(target).x
                y = MapNpc(mapnum).NPC(target).y
            End If
               
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Objetivo fuera de rango.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
           
            Select Case Spell(spellnum).Type
                Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_BUFF 'EaSee 0.5
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, Vital, spellnum
                                DidCast = True
                            End If
                        End If
                    End If
                   
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        increment = False
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    ElseIf Spell(spellnum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        increment = True
                        DidCast = True ' <--- Fixed!
                    End If
                   
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, target, True) Then
                                SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, increment, target, Vital, spellnum
                        End If
                    Else
                        If Spell(spellnum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, target, True) Then
                                SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapnum
                            End If
                        Else
                            SpellNpc_Effect VitalType, increment, target, Vital, spellnum, mapnum
                        End If
                    End If
                    
        
            End Select
    End Select
   
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
       
        TempPlayer(index).SpellCD(spellslot) = GetTickCount + (Spell(spellnum).CDTime * 1000)
        Call SendCooldown(index, spellslot)
        SendActionMsg mapnum, Trim$(Spell(spellnum).Name) & ".", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
    End If
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' send the sound
        SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, spellnum
        
        If increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Player index, spellnum
            End If
        ElseIf Not increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal spellnum As Long, ByVal mapnum As Long)
Dim sSymbol As String * 1
Dim Colour As Long

    If Damage > 0 Then
        If increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Colour = BrightGreen
            If Vital = Vitals.MP Then Colour = BrightBlue
        Else
            sSymbol = "-"
            Colour = Blue
        End If
    
        SendAnimation mapnum, Spell(spellnum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg mapnum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
        
        ' send the sound
        SendMapSound index, MapNpc(mapnum).NPC(index).x, MapNpc(mapnum).NPC(index).y, SoundEntity.seSpell, spellnum
        
        If increment Then
            If MapNpc(mapnum).NPC(index).Vital(Vital) + Damage <= GetNpcMaxVital(index, Vitals.HP) Then
                MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) + Damage
            Else
                MapNpc(mapnum).NPC(index).Vital(Vital) = GetNpcMaxVital(index, Vitals.HP)
            End If
            
            If Spell(spellnum).Duration > 0 Then
                AddHoT_Npc mapnum, index, spellnum
            End If
        ElseIf Not increment Then
            MapNpc(mapnum).NPC(index).Vital(Vital) = MapNpc(mapnum).NPC(index).Vital(Vital) - Damage
        End If
        ' send update
        SendMapNpcVitals mapnum, index
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(index).DoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal spellnum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(index).HoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long, ByVal Caster As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).DoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal spellnum As Long)
Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNpc(mapnum).NPC(index).HoT(I)
            If .Spell = spellnum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = spellnum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).Vital
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                   SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Spell(.Spell).Vital
                   Call SendVital(index, HP)
                Else
                   SendActionMsg Player(index).Map, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, Player(index).x * 32, Player(index).y * 32
                   SetPlayerVital index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Spell(.Spell).Vital
                   Call SendVital(index, HP)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(mapnum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, Spell(.Spell).Vital, , True
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal mapnum As Long, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(mapnum).NPC(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' time to tick?
            If GetTickCount > .Timer + (Spell(.Spell).Interval * 1000) Then
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                    MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = MapNpc(mapnum).NPC(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                    If MapNpc(mapnum).NPC(index).Vital(Vitals.HP) > GetNpcMaxVital(index, Vitals.HP) Then
                        MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = GetNpcMaxVital(index, Vitals.HP)
                    End If
                    Call SendVital(index, HP)
                Else
                    SendActionMsg mapnum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNpc(mapnum).NPC(index).x * 32, MapNpc(mapnum).NPC(index).y * 32
                    MapNpc(mapnum).NPC(index).Vital(Vitals.MP) = MapNpc(mapnum).NPC(index).Vital(Vitals.MP) + Spell(.Spell).Vital
                    If MapNpc(mapnum).NPC(index).Vital(Vitals.HP) > GetNpcMaxVital(index, Vitals.HP) Then
                        MapNpc(mapnum).NPC(index).Vital(Vitals.HP) = GetNpcMaxVital(index, Vitals.HP)
                    End If
                    Call SendVital(index, MP)
                End If
                .Timer = GetTickCount
                ' check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(spellnum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M22")
        PlayerMsg index, var1, BrightRed
    End If              '[COMBAT]M22
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal mapnum As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(mapnum).NPC(index).StunDuration = Spell(spellnum).StunDuration
        MapNpc(mapnum).NPC(index).StunTimer = GetTickCount
    End If
End Sub


Public Sub Confusion(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    If Spell(spellnum).Inversion > 0 Then
        ' set the values on index
        TempPlayer(index).ConfusionDuracion = Spell(spellnum).Duration
        TempPlayer(index).ConfusionTiempo = GetTickCount
        ' send it to the index
        EnviarConfusion index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M23")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M23
End Sub

Public Sub Paralizar(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Paralisis > 0 Then
        ' set the values on index
        TempPlayer(index).ParalisisDuracion = Spell(spellnum).Duration
        TempPlayer(index).ParalisisTiempo = GetTickCount
        ' send it to the index
        EnviarParalisis index
        EnviarSpriteColorH index, 2, TempPlayer(index).ParalisisDuracion
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M24")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M24
End Sub

Public Sub Envenenar(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Veneno > 0 Then
        ' set the values on index
        TempPlayer(index).VenenoDuracion = Spell(spellnum).Duration
        TempPlayer(index).VenenoGolpe = Spell(spellnum).VenenoDmg
        TempPlayer(index).VenenoTiempo = GetTickCount
        ' send it to the index
        EnviarVeneno index
        EnviarSpriteColorH index, 1, TempPlayer(index).VenenoDuracion
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M25")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M25
End Sub

Public Sub Velocidad(ByVal index As Long, ByVal spellnum As Long) 'EaSee 0.5
    ' check if it's a stunning spell
        ' set the values on index
        Dim var1 As String
        
        TempPlayer(index).VelocidadDuracion = Spell(spellnum).Duration
        TempPlayer(index).VelocidadCaminar2 = Spell(spellnum).VelocidadCaminar2
        TempPlayer(index).VelocidadCorrer2 = Spell(spellnum).VelocidadCorrer2
        TempPlayer(index).VelocidadBuff = Spell(spellnum).Velocidad
        TempPlayer(index).VelocidadTiempo = GetTickCount
       
        ' send it to the index
        EnviarVelocidad index
        ' tell him he's stunned
        PlayerMsg index, GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M26"), BrightRed
                            '[COMBAT]M26
End Sub

Public Sub Fuerza(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    If Spell(spellnum).Fuerza > 0 Then
        ' set the values on index
        TempPlayer(index).FuerzaDuracion = Spell(spellnum).Duration
        TempPlayer(index).FuerzaH = Spell(spellnum).Fuerza
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).FuerzaTiempo = GetTickCount
       
        ' send it to the index
        EnviarFuerzaH index
        ' tell him he's stunned
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M27")
        PlayerMsg index, var1, BrightRed
    End If              '[COMBAT]M27
End Sub

Public Sub Destreza(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Destreza > 0 Then
        ' set the values on index
        TempPlayer(index).DestrezaDuracion = Spell(spellnum).Duration
        TempPlayer(index).DestrezaH = Spell(spellnum).Destreza
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).DestrezaTiempo = GetTickCount
       
        ' send it to the index
        EnviarDestrezaH index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M28")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M28
End Sub

Public Sub Agilidad(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Agilidad > 0 Then
        ' set the values on index
        TempPlayer(index).AgilidadDuracion = Spell(spellnum).Duration
        TempPlayer(index).AgilidadH = Spell(spellnum).Agilidad
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).AgilidadTiempo = GetTickCount
       
        ' send it to the index
        EnviarAgilidadH index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M29")
        PlayerMsg index, var1, BrightRed
    End If              '[COMBAT]M29
End Sub

Public Sub Inteligencia(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Inteligencia > 0 Then
        ' set the values on index
        TempPlayer(index).InteligenciaDuracion = Spell(spellnum).Duration
        TempPlayer(index).InteligenciaH = Spell(spellnum).Inteligencia
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).InteligenciaTiempo = GetTickCount
       
        ' send it to the index
        EnviarInteligenciaH index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M30")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M30
End Sub

Public Sub Voluntad(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Voluntad > 0 Then
        ' set the values on index
        TempPlayer(index).VoluntadDuracion = Spell(spellnum).Duration
        TempPlayer(index).VoluntadH = Spell(spellnum).Voluntad
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).VoluntadTiempo = GetTickCount
       
        ' send it to the index
        EnviarVoluntadH index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M31")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M31
End Sub

Public Sub Sprite(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Sprite > 0 Then
        ' set the values on index
        TempPlayer(index).SpriteDuracion = Spell(spellnum).Duration
        Player(index).SpriteOR = Player(index).Sprite
        TempPlayer(index).SpriteNumero = Spell(spellnum).NumeroSprite
        TempPlayer(index).BuffValor = Spell(spellnum).Buff
        TempPlayer(index).SpriteTiempo = GetTickCount
       
        ' send it to the index
        EnviarSprite index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M32")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M32
End Sub

Public Sub Invisibilidad(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    Dim var1 As String
    
    If Spell(spellnum).Invisibilidad > 0 Then
        ' set the values on index
        TempPlayer(index).InvisibilidadDuracion = Spell(spellnum).Duration
        TempPlayer(index).InvisibilidadTiempo = GetTickCount
        ' send it to the index
        EnviarInvisibilidad index
        ' tell him he's stunned
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "COMBAT", "M33")
        PlayerMsg index, var1, BrightRed
    End If                  '[COMBAT]M33
End Sub

Public Sub TransportarH(ByVal index As Long, ByVal spellnum As Long)
    ' check if it's a stunning spell
    If Spell(spellnum).Transportar > 0 Then
        ' set the values on index
        TempPlayer(index).TransMapaH = Spell(spellnum).TransportarMapa
        TempPlayer(index).TransMapaxH = Spell(spellnum).TransportarX
        TempPlayer(index).TransMapayH = Spell(spellnum).TransportarY
        ' send it to the index
        EnviarTransportarH index
        ' tell him he's stunned
        End If
End Sub

