Attribute VB_Name = "modSvQuests"
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////
Option Explicit
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Constants
Public Const MAX_TASKS As Byte = 10
Public Const MAX_QUESTS As Byte = 70
Public Const MAX_QUESTS_ITEMS As Byte = 10 'Alatar v1.2

Public Const QUEST_TYPE_GOSLAY As Byte = 1
Public Const QUEST_TYPE_GOGATHER As Byte = 2
Public Const QUEST_TYPE_GOTALK As Byte = 3
Public Const QUEST_TYPE_GOREACH As Byte = 4
Public Const QUEST_TYPE_GOGIVE As Byte = 5
Public Const QUEST_TYPE_GOKILL As Byte = 6
Public Const QUEST_TYPE_GOTRAIN As Byte = 7
Public Const QUEST_TYPE_GOGET As Byte = 8
Public Const QUEST_TYPE_GOGETFROMEVENT As Byte = 9

Public Const QUEST_NOT_STARTED As Byte = 0
Public Const QUEST_STARTED As Byte = 1
Public Const QUEST_COMPLETED As Byte = 2
Public Const QUEST_COMPLETED_BUT As Byte = 3

'Types
Public Quest(1 To MAX_QUESTS) As QuestRec

Public Type PlayerQuestRec
    Status As Long
    ActualTask As Long
    CurrentCount As Long 'Used to handle the Amount property
End Type

'Alatar v1.2
Private Type QuestRequiredItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestGiveItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestTakeItemRec
    Item As Long
    Value As Long
End Type

Private Type QuestRewardItemRec
    Item As Long
    Value As Long
End Type
'/Alatar v1.2

Public Type TaskRec
    Order As Long
    NPC As Long
    Item As Long
    Map As Long
    Resource As Long
    Amount As Long
    Speech As String * 300
    TaskLog As String * 300
    QuestEnd As Boolean
    Event As Long
End Type

Public Type QuestRec
    'Alatar v1.2
    Name As String * 30
    Repeat As Long
    QuestLog As String * 300
    Speech(1 To 3) As String * 300
    GiveItem(1 To MAX_QUESTS_ITEMS) As QuestGiveItemRec
    TakeItem(1 To MAX_QUESTS_ITEMS) As QuestTakeItemRec
    
    RequiredLevel As Long
    RequiredQuest As Long
    RequiredClass(1 To 5) As Long
    RequiredItem(1 To MAX_QUESTS_ITEMS) As QuestRequiredItemRec
    
    RewardExp As Long
    RewardItem(1 To MAX_QUESTS_ITEMS) As QuestRewardItemRec
    
    Task(1 To MAX_TASKS) As TaskRec
    '/Alatar v1.2
    
    '/escfoe2 :p
    Skill As Long
    SkillExp As Long
 
End Type

' //////////////
' // DATABASE //
' //////////////

Sub SaveQuests()
    Dim I As Long
    For I = 1 To MAX_QUESTS
        Call SaveQuest(I)
    Next
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim F As Long, I As Long
    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        'Alatar v1.2
        Put #F, , Quest(QuestNum).Name
        Put #F, , Quest(QuestNum).Repeat
        Put #F, , Quest(QuestNum).QuestLog
        For I = 1 To 3
            Put #F, , Quest(QuestNum).Speech(I)
        Next
        For I = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).GiveItem(I)
        Next
        For I = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).TakeItem(I)
        Next
        Put #F, , Quest(QuestNum).RequiredLevel
        Put #F, , Quest(QuestNum).RequiredQuest
        For I = 1 To 5
            Put #F, , Quest(QuestNum).RequiredClass(I)
        Next
        For I = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).RequiredItem(I)
        Next
        Put #F, , Quest(QuestNum).RewardExp
        For I = 1 To MAX_QUESTS_ITEMS
            Put #F, , Quest(QuestNum).RewardItem(I)
        Next
        For I = 1 To MAX_TASKS
            Put #F, , Quest(QuestNum).Task(I)
        Next
        '/Alatar v1.2
    Close #F
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim I As Integer
    Dim F As Long, n As Long
    Dim sLen As Long
    
    Call CheckQuests

    For I = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & I & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        
        'Alatar v1.2
        Get #F, , Quest(I).Name
        Get #F, , Quest(I).Repeat
        Get #F, , Quest(I).QuestLog
        For n = 1 To 3
            Get #F, , Quest(I).Speech(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).GiveItem(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).TakeItem(n)
        Next
        Get #F, , Quest(I).RequiredLevel
        Get #F, , Quest(I).RequiredQuest
        For n = 1 To 5
            Get #F, , Quest(I).RequiredClass(n)
        Next
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).RequiredItem(n)
        Next
        Get #F, , Quest(I).RewardExp
        For n = 1 To MAX_QUESTS_ITEMS
            Get #F, , Quest(I).RewardItem(n)
        Next
        For n = 1 To MAX_TASKS
            Get #F, , Quest(I).Task(n)
        Next
        '/Alatar v1.2
        Close #F
    Next
End Sub

Sub CheckQuests()
    Dim I As Long
    For I = 1 To MAX_QUESTS
        If Not FileExist("\Data\quests\quest" & I & ".dat") Then
            Call SaveQuest(I)
        End If
    Next
End Sub

Sub ClearQuest(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Quest(index).Name = vbNullString
    Quest(index).QuestLog = vbNullString
End Sub

Sub ClearQuests()
    Dim I As Long

    For I = 1 To MAX_QUESTS
        Call ClearQuest(I)
    Next
End Sub

' ////////////////////
' // C&S PROCEDURES //
' ////////////////////

Sub SendQuests(ByVal index As Long)
    Dim I As Long

    For I = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(I).Name)) > 0 Then
            Call SendUpdateQuestTo(index, I)
        End If
    Next
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong QuestNum
    buffer.WriteBytes QuestData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim packet As String
    Dim buffer As clsBuffer
    Dim QuestSize As Long
    Dim QuestData() As Byte
    Set buffer = New clsBuffer
    QuestSize = LenB(Quest(QuestNum))
    ReDim QuestData(QuestSize - 1)
    CopyMemory QuestData(0), ByVal VarPtr(Quest(QuestNum)), QuestSize
    buffer.WriteLong SUpdateQuest
    buffer.WriteLong QuestNum
    buffer.WriteBytes QuestData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendPlayerQuests(ByVal index As Long)
    Dim I As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
        For I = 1 To MAX_QUESTS
            buffer.WriteLong Player(index).PlayerQuest(I).Status
            buffer.WriteLong Player(index).PlayerQuest(I).ActualTask
            buffer.WriteLong Player(index).PlayerQuest(I).CurrentCount
        Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

End Sub

Public Sub SendPlayerQuest(ByVal index As Long, ByVal QuestNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerQuest
    buffer.WriteLong Player(index).PlayerQuest(QuestNum).Status
    buffer.WriteLong Player(index).PlayerQuest(QuestNum).ActualTask
    buffer.WriteLong Player(index).PlayerQuest(QuestNum).CurrentCount
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

'sends a message to the client that is shown on the screen
Public Sub QuestMessage(ByVal index As Long, ByVal QuestNum As Long, ByVal message As String, ByVal QuestNumForStart As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SQuestMessage
    buffer.WriteLong QuestNum
    buffer.WriteString Trim$(message)
    buffer.WriteLong QuestNumForStart
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
End Sub

' ///////////////
' // Functions //
' ///////////////

Public Function CanStartQuest(ByVal index As Long, ByVal QuestNum As Long) As Boolean
    Dim I As Long, n As Long
    CanStartQuest = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    If QuestInProgress(index, QuestNum) Then Exit Function
    
    'check if now a completed quest can be repeated
    If Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Then
        If Quest(QuestNum).Repeat = YES Then
            Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT
            Exit Function
        End If
    End If
    
    'Check if player has the quest 0 (not started) or 3 (completed but it can be started again)
    If Player(index).PlayerQuest(QuestNum).Status = QUEST_NOT_STARTED Or Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        'Check if player's level is right
        If Quest(QuestNum).RequiredLevel <= Player(index).Level Then
            
            'Check if item is needed
            For I = 1 To MAX_QUESTS_ITEMS
                If Quest(QuestNum).RequiredItem(I).Item > 0 Then
                    'if we don't have it at all then
                    If HasItem(index, Quest(QuestNum).RequiredItem(I).Item) = 0 Then
                        PlayerMsg index, "Necesitas " & CheckGrammar(Trim$(Item(Quest(QuestNum).RequiredItem(I).Item).Name), 0) & " para comenzar esta Mision.", BrightRed
                        Exit Function
                    End If
                End If
            Next
            
            'Check if previous quest is needed
            If Quest(QuestNum).RequiredQuest > 0 And Quest(QuestNum).RequiredQuest <= MAX_QUESTS Then
                If Player(index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_NOT_STARTED Or Player(index).PlayerQuest(Quest(QuestNum).RequiredQuest).Status = QUEST_STARTED Then
                    PlayerMsg index, "Necesitas completar " & Trim$(Quest(Quest(QuestNum).RequiredQuest).Name) & " primero para empezar esta Mision.", BrightRed
                    Exit Function
                End If
            End If
            'Go on :)
            CanStartQuest = True
        Else
            PlayerMsg index, "Requieres nivel " & Quest(QuestNum).RequiredLevel & " para comenzar esta Mision.", BrightRed
        End If
    Else
        PlayerMsg index, "No puedes empezar esta Mision nuevamente!", BrightRed
    End If
End Function

Public Function CanEndQuest(ByVal index As Long, QuestNum As Long) As Boolean
    CanEndQuest = False
    If Quest(QuestNum).Task(Player(index).PlayerQuest(QuestNum).ActualTask).QuestEnd = True Then
        CanEndQuest = True
    End If
End Function

'Tells if the quest is in progress or not
Public Function QuestInProgress(ByVal index As Long, ByVal QuestNum As Long) As Boolean
    QuestInProgress = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(index).PlayerQuest(QuestNum).Status = QUEST_STARTED Then
        QuestInProgress = True
    End If
End Function

Public Function QuestCompleted(ByVal index As Long, ByVal QuestNum As Long) As Boolean
    QuestCompleted = False
    If QuestNum < 1 Or QuestNum > MAX_QUESTS Then Exit Function
    
    If Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED Or Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED_BUT Then
        QuestCompleted = True
    End If
End Function

'Gets the quest reference num (id) from the quest name (it shall be unique)
Public Function GetQuestNum(ByVal QuestName As String) As Long
    Dim I As Long
    GetQuestNum = 0
    
    For I = 1 To MAX_QUESTS
        If Trim$(Quest(I).Name) = Trim$(QuestName) Then
            GetQuestNum = I
            Exit For
        End If
    Next
End Function

Public Function GetItemNum(ByVal ItemName As String) As Long
    Dim I As Long
    GetItemNum = 0
    
    For I = 1 To MAX_ITEMS
        If Trim$(Item(I).Name) = Trim$(ItemName) Then
            GetItemNum = I
            Exit For
        End If
    Next
End Function

' /////////////////////
' // General Purpose //
' /////////////////////

Public Sub CheckTasks(ByVal index As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim I As Long
    
    For I = 1 To MAX_QUESTS
        If QuestInProgress(index, I) Then
            If TaskType = Quest(I).Task(Player(index).PlayerQuest(I).ActualTask).Order Then
                Call CheckTask(index, I, TaskType, TargetIndex)
            End If
        End If
    Next
End Sub

Public Sub CheckTask(ByVal index As Long, ByVal QuestNum As Long, ByVal TaskType As Long, ByVal TargetIndex As Long)
    Dim ActualTask As Long, I As Long
    ActualTask = Player(index).PlayerQuest(QuestNum).ActualTask
    
    Select Case TaskType
        Case QUEST_TYPE_GOSLAY 'Kill X amount of X npc's.
        
            'is npc's defeated id is the same as the npc i have to kill?
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                'Count +1
                Player(index).PlayerQuest(QuestNum).CurrentCount = Player(index).PlayerQuest(QuestNum).CurrentCount + 1
                'show msg
                PlayerMsg index, "Mision: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(NPC(TargetIndex).Name) + " asesinado.", Yellow
                'did i finish the work?
                If Player(index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage index, QuestNum, "Tarea Completada", 0
                    'is the quest's end?
                    If CanEndQuest(index, QuestNum) Then
                        EndQuest index, QuestNum
                    Else
                        'otherwise continue to the next task
                        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                        
        Case QUEST_TYPE_GOGATHER 'Gather X amount of X item.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Item Then
                
                'reset the count first
                Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                
                'Check inventory for the items
                For I = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, I) = TargetIndex Then
                        If Item(I).Type = ITEM_TYPE_CURRENCY Then
                            Player(index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(index, I)
                        Else
                            'If is the correct item add it to the count
                            Player(index).PlayerQuest(QuestNum).CurrentCount = Player(index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                PlayerMsg index, "Mision: " + Trim$(Quest(QuestNum).Name) + " - Tienes " + Trim$(Player(index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                
                If Player(index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage index, QuestNum, "Tarea Completada", 0
                    If CanEndQuest(index, QuestNum) Then
                        EndQuest index, QuestNum
                    Else
                        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
            
        Case QUEST_TYPE_GOTALK 'Interact with X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                QuestMessage index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(index, QuestNum) Then
                    EndQuest index, QuestNum
                Else
                    Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOREACH 'Reach X map.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Map Then
                QuestMessage index, QuestNum, "Tarea Completada", 0
                If CanEndQuest(index, QuestNum) Then
                    EndQuest index, QuestNum
                Else
                    Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        
        Case QUEST_TYPE_GOGIVE 'Give X amount of X item to X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                
                Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                
                For I = 1 To MAX_INV
                    If GetPlayerInvItemNum(index, I) = Quest(QuestNum).Task(ActualTask).Item Then
                        If Item(I).Type = ITEM_TYPE_CURRENCY Then
                            If GetPlayerInvItemValue(index, I) >= Quest(QuestNum).Task(ActualTask).Amount Then
                                Player(index).PlayerQuest(QuestNum).CurrentCount = GetPlayerInvItemValue(index, I)
                            End If
                        Else
                            'If is the correct item add it to the count
                            Player(index).PlayerQuest(QuestNum).CurrentCount = Player(index).PlayerQuest(QuestNum).CurrentCount + 1
                        End If
                    End If
                Next
                
                If Player(index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    'if we have enough items, then remove them and finish the task
                    If Item(Quest(QuestNum).Task(ActualTask).Item).Type = ITEM_TYPE_CURRENCY Then
                        TakeInvItem index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                    Else
                        'If it's not a currency then remove all the items
                        For I = 1 To Quest(QuestNum).Task(ActualTask).Amount
                            TakeInvItem index, Quest(QuestNum).Task(ActualTask).Item, 1
                        Next
                    End If
                    
                    PlayerMsg index, "Mision: " + Trim$(Quest(QuestNum).Name) + " - Le diste " + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " " + Trim$(Item(TargetIndex).Name), Yellow
                    QuestMessage index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                    
                    If CanEndQuest(index, QuestNum) Then
                        EndQuest index, QuestNum
                    Else
                        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                    
        Case QUEST_TYPE_GOKILL 'Kill X amount of players.
            Player(index).PlayerQuest(QuestNum).CurrentCount = Player(index).PlayerQuest(QuestNum).CurrentCount + 1
            PlayerMsg index, "Mision: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " jugadores aniquilados.", Yellow
            If Player(index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                QuestMessage index, QuestNum, "Tarea Completada", 0
                If CanEndQuest(index, QuestNum) Then
                    EndQuest index, QuestNum
                Else
                    Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                    Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
            
        Case QUEST_TYPE_GOTRAIN 'Hit X amount of times X resource.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).Resource Then
                Player(index).PlayerQuest(QuestNum).CurrentCount = Player(index).PlayerQuest(QuestNum).CurrentCount + 1
                PlayerMsg index, "Mision: " + Trim$(Quest(QuestNum).Name) + " - " + Trim$(Player(index).PlayerQuest(QuestNum).CurrentCount) + "/" + Trim$(Quest(QuestNum).Task(ActualTask).Amount) + " golpes.", Yellow
                If Player(index).PlayerQuest(QuestNum).CurrentCount >= Quest(QuestNum).Task(ActualTask).Amount Then
                    QuestMessage index, QuestNum, "Tarea Completada", 0
                    If CanEndQuest(index, QuestNum) Then
                        EndQuest index, QuestNum
                    Else
                        Player(index).PlayerQuest(QuestNum).CurrentCount = 0
                        Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                    End If
                End If
            End If
                      
        Case QUEST_TYPE_GOGET 'Get X amount of X item from X npc.
            If TargetIndex = Quest(QuestNum).Task(ActualTask).NPC Then
                GiveInvItem index, Quest(QuestNum).Task(ActualTask).Item, Quest(QuestNum).Task(ActualTask).Amount
                QuestMessage index, QuestNum, Quest(QuestNum).Task(ActualTask).Speech, 0
                If CanEndQuest(index, QuestNum) Then
                    EndQuest index, QuestNum
                Else
                    Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
                End If
            End If
        'For when I feel like adding it...
        'Case QUEST_TYPE_GOGETFROMEVENT 'Get x amount of item x from npc x
        '    If TargetIndex = 3 Then 'Quest(QuestNum).Task(ActualTask).Event Then
        '        Call PlayerMsg(index, "Well we got this far... CheckTask-modSvQuests", BrightRed)
        '    End If
        '
        '    If CanEndQuest(index, QuestNum) Then
        '        EndQuest index, QuestNum
        '    Else
        '        Player(index).PlayerQuest(QuestNum).ActualTask = ActualTask + 1
        '    End If
        
    End Select
    SavePlayer index
    SendPlayerData index
    SendPlayerQuests index
End Sub

Public Sub EndQuest(ByVal index As Long, ByVal QuestNum As Long)
    Dim I As Long, n As Long
    
    Player(index).PlayerQuest(QuestNum).Status = QUEST_COMPLETED
    
    'reset counters to 0
    Player(index).PlayerQuest(QuestNum).ActualTask = 0
    Player(index).PlayerQuest(QuestNum).CurrentCount = 0
    
    'give experience
    GivePlayerEXP index, Quest(QuestNum).RewardExp
    
    'give skill experience
    If Quest(QuestNum).Skill > 0 Then
        Call SetPlayerSkillExp(index, Quest(QuestNum).Skill, Quest(QuestNum).SkillExp)
    End If
    
    'remove items on the end
    For I = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).TakeItem(I).Item > 0 Then
            If HasItem(index, Quest(QuestNum).TakeItem(I).Item) > 0 Then
                If Item(Quest(QuestNum).TakeItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                    TakeInvItem index, Quest(QuestNum).TakeItem(I).Item, Quest(QuestNum).TakeItem(I).Value
                Else
                    For n = 1 To Quest(QuestNum).TakeItem(I).Value
                        TakeInvItem index, Quest(QuestNum).TakeItem(I).Item, 1
                    Next
                End If
            End If
        End If
    Next
        
    SavePlayer index
    Call SendStats(index)
    SendPlayerData index
    
    'give rewards
    For I = 1 To MAX_QUESTS_ITEMS
        If Quest(QuestNum).RewardItem(I).Item <> 0 Then
            'check if we have space
            If FindOpenInvSlot(index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                PlayerMsg index, "Necesitas espacio en tu Inventario.", BrightRed
                Exit For
            Else
                'if so, check if it's a currency stack the item in one slot
                If Item(Quest(QuestNum).RewardItem(I).Item).Type = ITEM_TYPE_CURRENCY Then
                    GiveInvItem index, Quest(QuestNum).RewardItem(I).Item, Quest(QuestNum).RewardItem(I).Value
                Else
                'if not, create a new loop and store the item in a new slot if is possible
                    For n = 1 To Quest(QuestNum).RewardItem(I).Value
                        If FindOpenInvSlot(index, Quest(QuestNum).RewardItem(I).Item) = 0 Then
                            PlayerMsg index, "Necesitas espacio en tu Inventario.", BrightRed
                            Exit For
                        Else
                            GiveInvItem index, Quest(QuestNum).RewardItem(I).Item, 1
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    'show ending message
    QuestMessage index, QuestNum, Trim$(Quest(QuestNum).Speech(3)), 0
    
    'mark quest as completed in chat
    PlayerMsg index, Trim$(Quest(QuestNum).Name) & ": Mision Completada", Green
    
    SavePlayer index
    SendEXP index
    Call SendStats(index)
    SendPlayerData index
    SendPlayerQuests index
End Sub
