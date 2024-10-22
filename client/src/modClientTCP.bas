Attribute VB_Name = "modClientTCP"
Option Explicit
'-------------------SINGLEPLAYER
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (lpidProcess As Long, ByVal cb As Long, cbNeeded As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, lphModule As Long, ByVal cb As Long, lpcbNeeded As Long) As Long
Private Declare Function GetModuleBaseName Lib "PSAPI.DLL" Alias "GetModuleBaseNameA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As String, ByVal nsize As Long) As Long
Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_QUERY_INFORMATION = &H400
' ******************************************
' ** Communcation to server, TCP          **
' ** Winsock Control (mswinsck.ocx)       **
' ** String packets (slow and big)        **
' ******************************************
Private PlayerBuffer As clsBuffer
Public MDE As String
Dim PAC As String



Sub TcpInit()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set PlayerBuffer = New clsBuffer

    ' connect
    If MDE = 1 Then
    frmMain.Socket.RemoteHost = frmMenu.IPSINGLE
    frmMain.Socket.RemotePort = frmMenu.PORTSINGLE
    Else
    frmMain.Socket.RemoteHost = Options.IP
    frmMain.Socket.RemotePort = Options.Port
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TcpInit", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DestroyTCP()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.Socket.Close
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DestroyTCP", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    frmMain.Socket.GetData buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes buffer()
    
    If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.Length - 4
        If pLength <= PlayerBuffer.Length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.Length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "IncomingData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConnectToServer(ByVal I As Long) As Boolean
Dim Wait As Long
Dim CSW As String
Dim IAA As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    CSW = 0
    If MDE = "1" Then
    If frmMenu.IPSINGLE = "127.0.0.1" Then
    Call PutVar(App.Path & "\data files\motor\data\options.ini", "OPTIONS", "Port", frmMenu.PORTSINGLE)
    If EstaCorriendo("Motor.exe") = False Then Shell App.Path & "\data files\motor\Motor.exe -#Mode$Singer$#"
    End If
    End If
LoadC:
    ' Check to see if we are already connected, if so just exit
    If MDE = "1" Then
    If frmMenu.IPSINGLE = "127.0.0.1" Then CSW = 1
    End If
    
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    'MsgBox GetTickCount
    frmMain.Socket.Close
    frmMain.Socket.Connect
    
    If MDE = "1" Then
    SetStatus "Entrando a la partida..."
    Else
    SetStatus "Conectando al Servidor..."
    End If
    
    ' Wait until connected or 3 seconds have passed and report the server being down
    If CSW = 1 Then
     Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
           DoEvents
         Loop
          If IsConnected = False Then GoTo LoadC
          ConnectToServer = IsConnected
    Else
     'Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        'DoEvents
     'Loop
        ConnectToServer = IsConnected
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "ConnectToServer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsConnected() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmMain.Socket.state = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsConnected", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' if the player doesn't exist, the name will equal 0
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsPlaying", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
Sub SendData(ByRef data() As Byte)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsConnected Then
        Set buffer = New clsBuffer
                
        buffer.WriteLong (UBound(data) - LBound(data)) + 1
        buffer.WriteBytes data()
        frmMain.Socket.SendData buffer.ToArray()
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    
    Exit Sub
End Sub

' *****************************
' ** Outgoing Client Packets **
' *****************************
Public Sub SendNewAccount(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CNewAccount
    buffer.WriteString name
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendNewAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDblClickPos()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CClickPos
    buffer.WriteLong myTarget
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendClickPos", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
    
End Sub


Public Sub SendDelAccount(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDelAccount
    buffer.WriteString name
    buffer.WriteString Password
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDelAccount", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal name As String, ByVal Password As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'MsgBox ("1")
    Set buffer = New clsBuffer
    buffer.WriteLong CLogin
    buffer.WriteString name
    buffer.WriteString Password
    buffer.WriteLong "2"
    buffer.WriteLong App.Minor
    buffer.WriteLong App.Revision
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendLogin", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal Sprite As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAddChar
    buffer.WriteString name
    buffer.WriteLong Sex
    buffer.WriteLong ClassNum
    buffer.WriteLong Sprite
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAddChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseChar
    buffer.WriteLong CharSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUseChar", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendFollowPlayer(ByVal Who As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CFollowPlayer
    buffer.WriteLong Who
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendFollowPlayer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SayMsg(ByVal text As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSayMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub OrderAchievements(ByVal x As String)
Dim buffer As clsBuffer
    
    x = "1"
    ' If debug mode, handle error then exit out
    
    Set buffer = New clsBuffer
    buffer.WriteLong COrderAchievements
    buffer.WriteString x
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
End Sub
Public Sub SendCommand(ByVal Comando As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSendCommandForClient
    buffer.WriteString Comando
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SayMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BroadcastMsg(ByVal text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBroadcastMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "BroadcastMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EmoteMsg(ByVal text As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEmoteMsg
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "EmoteMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RequestFriendData(ByVal fName As String)
Dim buffer As clsBuffer
    Set buffer = New clsBuffer
        
    buffer.WriteLong CRequestFriendData
    buffer.WriteString fName
    
    SendData buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub PrivateMsg(ByVal text As String, ByVal MsgTo As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPrivateMsg
    buffer.WriteString MsgTo
    buffer.WriteString text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PrivateMsg", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerMove
    buffer.WriteLong GetPlayerDir(MyIndex)
    buffer.WriteLong Player(MyIndex).Moving
    buffer.WriteLong Player(MyIndex).x
    buffer.WriteLong Player(MyIndex).y
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerMove", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerDir
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerDir", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNewMap
    buffer.WriteLong GetPlayerDir(MyIndex)
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPlayerRequestNewMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim packet As String
Dim x As Long
Dim y As Long
Dim I As Long, Z As Long, w As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    CanMoveNow = False
    With Map
        buffer.WriteLong CMapData
        buffer.WriteString Trim$(.name)
        buffer.WriteString Trim$(.Music)
        buffer.WriteString Trim$(.BGS)
        buffer.WriteByte .Moral
        buffer.WriteLong .Up
        buffer.WriteLong .Down
        buffer.WriteLong .Left
        buffer.WriteLong .Right
        buffer.WriteLong .BootMap
        buffer.WriteByte .BootX
        buffer.WriteByte .BootY
        
        buffer.WriteLong Map.Weather
        buffer.WriteLong Map.WeatherIntensity
        
        buffer.WriteLong Map.Fog
        buffer.WriteLong Map.FogSpeed
        buffer.WriteLong Map.FogOpacity
        
        buffer.WriteLong Map.Red
        buffer.WriteLong Map.Green
        buffer.WriteLong Map.Blue
        buffer.WriteLong Map.alpha
        
        buffer.WriteByte .MaxX
        buffer.WriteByte .MaxY
        
        buffer.WriteByte .DropItemsOnDeath
    End With

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY

            With Map.Tile(x, y)
                For I = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .layer(I).x
                    buffer.WriteLong .layer(I).y
                    buffer.WriteLong .layer(I).Tileset
                Next
                For Z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(Z)
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

    With Map

        For x = 1 To MAX_MAP_NPCS
            buffer.WriteLong .NPC(x)
            buffer.WriteLong .NpcSpawnType(x)
        Next

    End With
    
    'Event Data
    buffer.WriteLong Map.EventCount
        
    If Map.EventCount > 0 Then
        For I = 1 To Map.EventCount
            With Map.Events(I)
                buffer.WriteString .name
                buffer.WriteLong .Global
                buffer.WriteLong .x
                buffer.WriteLong .y
                buffer.WriteLong .pageCount
            End With
            If Map.Events(I).pageCount > 0 Then
                For x = 1 To Map.Events(I).pageCount
                    With Map.Events(I).Pages(x)
                        buffer.WriteLong .chkVariable
                        buffer.WriteLong .VariableIndex
                        buffer.WriteLong .VariableCondition
                        buffer.WriteLong .VariableCompare
                            
                        buffer.WriteLong .chkSwitch
                        buffer.WriteLong .SwitchIndex
                        buffer.WriteLong .SwitchCompare
                        
                        buffer.WriteLong .chkHasItem
                        buffer.WriteLong .HasItemIndex
                        buffer.WriteLong .HasItemAmount
                            
                        buffer.WriteLong .chkSelfSwitch
                        buffer.WriteLong .SelfSwitchIndex
                        buffer.WriteLong .SelfSwitchCompare
                            
                        buffer.WriteLong .GraphicType
                        buffer.WriteLong .Graphic
                        buffer.WriteLong .GraphicX
                        buffer.WriteLong .GraphicY
                        buffer.WriteLong .GraphicX2
                        buffer.WriteLong .GraphicY2
                        
                        buffer.WriteLong .MoveType
                        buffer.WriteLong .MoveSpeed
                        buffer.WriteLong .MoveFreq
                        buffer.WriteLong .MoveRouteCount
                        
                        buffer.WriteLong .IgnoreMoveRoute
                        buffer.WriteLong .RepeatMoveRoute
                            
                        If .MoveRouteCount > 0 Then
                            For y = 1 To .MoveRouteCount
                                buffer.WriteLong .MoveRoute(y).Index
                                buffer.WriteLong .MoveRoute(y).Data1
                                buffer.WriteLong .MoveRoute(y).Data2
                                buffer.WriteLong .MoveRoute(y).Data3
                                buffer.WriteLong .MoveRoute(y).Data4
                                buffer.WriteLong .MoveRoute(y).Data5
                                buffer.WriteLong .MoveRoute(y).Data6
                            Next
                        End If
                            
                        buffer.WriteLong .WalkAnim
                        buffer.WriteLong .DirFix
                        buffer.WriteLong .Walkthrough
                        buffer.WriteLong .ShowName
                        buffer.WriteLong .Trigger
                        buffer.WriteLong .CommandListCount
                        
                        buffer.WriteLong .Position
                    End With
                        
                    If Map.Events(I).Pages(x).CommandListCount > 0 Then
                        For y = 1 To Map.Events(I).Pages(x).CommandListCount
                            buffer.WriteLong Map.Events(I).Pages(x).CommandList(y).CommandCount
                            buffer.WriteLong Map.Events(I).Pages(x).CommandList(y).ParentList
                            If Map.Events(I).Pages(x).CommandList(y).CommandCount > 0 Then
                                For Z = 1 To Map.Events(I).Pages(x).CommandList(y).CommandCount
                                    With Map.Events(I).Pages(x).CommandList(y).Commands(Z)
                                        buffer.WriteLong .Index
                                        buffer.WriteString .Text1
                                        buffer.WriteString .Text2
                                        buffer.WriteString .Text3
                                        buffer.WriteString .Text4
                                        buffer.WriteString .Text5
                                        buffer.WriteLong .Data1
                                        buffer.WriteLong .Data2
                                        buffer.WriteLong .Data3
                                        buffer.WriteLong .Data4
                                        buffer.WriteLong .Data5
                                        buffer.WriteLong .Data6
                                        buffer.WriteLong .ConditionalBranch.CommandList
                                        buffer.WriteLong .ConditionalBranch.Condition
                                        buffer.WriteLong .ConditionalBranch.Data1
                                        buffer.WriteLong .ConditionalBranch.Data2
                                        buffer.WriteLong .ConditionalBranch.Data3
                                        buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        buffer.WriteLong .MoveRouteCount
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                buffer.WriteLong .MoveRoute(w).Index
                                                buffer.WriteLong .MoveRoute(w).Data1
                                                buffer.WriteLong .MoveRoute(w).Data2
                                                buffer.WriteLong .MoveRoute(w).Data3
                                                buffer.WriteLong .MoveRoute(w).Data4
                                                buffer.WriteLong .MoveRoute(w).Data5
                                                buffer.WriteLong .MoveRoute(w).Data6
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

    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpMeTo(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpMeTo
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarpMeTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpToMe(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpToMe
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarptoMe", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WarpTo(ByVal MapNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWarpTo
    buffer.WriteLong MapNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WarpTo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetAccess
    buffer.WriteString name
    buffer.WriteLong Access
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetAccess", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetName(ByVal name As String, ByVal newName As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetName
    buffer.WriteString name
    buffer.WriteString newName
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetName", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHealPlayer(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CHealPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendHealPlayer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKillPlayer(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteLong CKillPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendKillPlayer", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSetSprite(ByVal SpriteNum As Long, ByVal plName As String)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetSprite
    buffer.WriteLong SpriteNum
    buffer.WriteString plName
    
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendKick(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CKickPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendKick", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWalkthrough()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWalkthrough
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendWalkthrough", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBan(ByVal name As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanPlayer
    buffer.WriteString name
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBan", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendOpenBankCommand()
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong COpenMyBank
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendOpenBankCommand", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanList()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanList
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBanList", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditItem
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditCombo()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditCombo
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditCombo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal itemNum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(itemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(itemNum)), ItemSize
    buffer.WriteLong CSaveItem
    buffer.WriteLong itemNum
    buffer.WriteBytes ItemData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveCombo(ByVal comboNum As Long)
Dim buffer As clsBuffer
Dim ComboSize As Long
Dim ComboData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    ComboSize = LenB(Combo(comboNum))
    ReDim ComboData(ComboSize - 1)
    CopyMemory ComboData(0), ByVal VarPtr(Combo(comboNum)), ComboSize
    buffer.WriteLong CSaveCombo
    buffer.WriteLong comboNum
    buffer.WriteBytes ComboData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveCombo", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDeleteFriend(ByVal fName As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CDeleteFriend
    buffer.WriteString fName
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeleteFriend", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditAnimation()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditAnimation
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveAnimation(ByVal Animationnum As Long)
Dim buffer As clsBuffer
Dim AnimationSize As Long
Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    AnimationSize = LenB(Animation(Animationnum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(Animationnum)), AnimationSize
    buffer.WriteLong CSaveAnimation
    buffer.WriteLong Animationnum
    buffer.WriteBytes AnimationData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveAnimation", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNpc()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditNpc
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNpc(ByVal npcNum As Long)
Dim buffer As clsBuffer
Dim NpcSize As Long
Dim NpcData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    NpcSize = LenB(NPC(npcNum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcNum)), NpcSize
    buffer.WriteLong CSaveNpc
    buffer.WriteLong npcNum
    buffer.WriteBytes NpcData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveNpc", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditResource()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditResource
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveResource(ByVal ResourceNum As Long)
Dim buffer As clsBuffer
Dim ResourceSize As Long
Dim ResourceData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    buffer.WriteLong CSaveResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveResource", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapRespawn()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapRespawn
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendCheckHighlightItem(ByVal invNum As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CHighlightItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SencCheckHighlightItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal invNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseItem
    buffer.WriteLong invNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUseItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAcceptFriend(ByVal fName As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CFriendAccept
    buffer.WriteString fName
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAcceptFriend", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDeclineFriend(ByVal fName As String)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CFriendDecline
    buffer.WriteString fName
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeclineFriend", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDropItem(ByVal invNum As Long, ByVal Amount As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InBank Or InShop Then Exit Sub
    
    ' do basic checks
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    If PlayerInv(invNum).num < 1 Or PlayerInv(invNum).num > MAX_ITEMS Then Exit Sub
    If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
        If Amount < 1 Or Amount > PlayerInv(invNum).Value Then Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapDropItem
    buffer.WriteLong invNum
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDropItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWhosOnline()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWhosOnline
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendWhosOnline", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSetMotd
    buffer.WriteString MOTD
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMOTDChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditShop
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal shopnum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(shopnum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(shopnum)), ShopSize
    buffer.WriteLong CSaveShop
    buffer.WriteLong shopnum
    buffer.WriteBytes ShopData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveShop", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditSpell()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditSpell
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveSpell(ByVal spellnum As Long)
Dim buffer As clsBuffer
Dim SpellSize As Long
Dim SpellData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    SpellSize = LenB(Spell(spellnum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(spellnum)), SpellSize
    
    buffer.WriteLong CSaveSpell
    buffer.WriteLong spellnum
    buffer.WriteBytes SpellData
    SendData buffer.ToArray()
    
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSaveSpell", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestEditMap
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestEditMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBanDestroy()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBanDestroy
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendBanDestroy", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeInvSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapInvSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendChangeSpellSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSwapSpellSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendChangeInvSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GetPing()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingStart = GetTickCount
    Set buffer = New clsBuffer
    buffer.WriteLong CCheckPing
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "GetPing", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUnequip(ByVal eqNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUnequip
    buffer.WriteLong eqNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendUnequip", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestPlayerData()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestPlayerData
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestPlayerData", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestItems()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestItems
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestItems", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestCombos()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestCombos
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestCombos", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestAnimations()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestAnimations
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestAnimations", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestNPCS()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestNPCS
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestNPCS", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestResources()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestResources
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestResources", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestSpells()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestSpells
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestSpells", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendRequestShops()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestShops
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestShops", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItem(ByVal tmpItem As Long, ByVal tmpAmount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSpawnItem
    buffer.WriteLong tmpItem
    buffer.WriteLong tmpAmount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSpawnItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrainStat(ByVal StatNum As Byte)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUseStatPoint
    buffer.WriteByte StatNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestLevelUp(ByVal PlayerName As String)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRequestLevelUp
    buffer.WriteString PlayerName
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendRequestLevelUp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub BuyItem(ByVal shopSlot As Long)
Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CBuyItem
    buffer.WriteLong shopSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "BuyItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SellItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CSellItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SellItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDepositItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DepositItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawItem(ByVal bankslot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CWithdrawItem
    buffer.WriteLong bankslot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "WithdrawItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBank()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCloseBank
    SendData buffer.ToArray()
    Set buffer = Nothing
    InBank = False
    GUIWindow(GUI_BANK).Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CloseBank", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChangeBankSlots(ByVal OldSlot As Long, ByVal NewSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CChangeBankSlots
    buffer.WriteLong OldSlot
    buffer.WriteLong NewSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ChangeBankSlots", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal x As Long, ByVal y As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAdminWarp
    buffer.WriteLong x
    buffer.WriteLong y
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AdminWarp", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AcceptTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "AcceptTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DeclineTrade()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTrade
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "DeclineTrade", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TradeItem(ByVal invSlot As Long, ByVal Amount As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeItem
    buffer.WriteLong invSlot
    buffer.WriteLong Amount
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "TradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UntradeItem(ByVal invSlot As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CUntradeItem
    buffer.WriteLong invSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "UntradeItem", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarChange(ByVal sType As Long, ByVal Slot As Long, ByVal hotbarNum As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarChange
    buffer.WriteLong sType
    buffer.WriteLong Slot
    buffer.WriteLong hotbarNum
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendHotbarChange", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHotbarUse(ByVal Slot As Long)
Dim buffer As clsBuffer, x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' check if spell
    If Hotbar(Slot).sType = 2 Then ' spell
        For x = 1 To MAX_PLAYER_SPELLS
            ' is the spell matching the hotbar?
            If PlayerSpells(x) = Hotbar(Slot).Slot Then
                ' found it, cast it
                CastSpell x
                Exit Sub
            End If
        Next
        ' can't find the spell, exit out
        Exit Sub
    End If

    Set buffer = New clsBuffer
    buffer.WriteLong CHotbarUse
    buffer.WriteLong Slot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendHotbarUse", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendVisibility()
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPlayerVisibility
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    If Not GetPlayerVisible(MyIndex) = 1 Then
    Call AddText("(Te vuelves invisible)", AlertColor)
    Else
    Call AddText("(Te vuelves visible)", AlertColor)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendVisibility", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapReport()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapReport
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMapReport", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSearch(ByVal CurX As Long, ByVal CurY As Long)
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If isInBounds Then
        Set buffer = New clsBuffer
        buffer.WriteLong CSearch
        buffer.WriteLong CurX
        buffer.WriteLong CurY
        SendData buffer.ToArray()
        Set buffer = Nothing
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PlayerSearch", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAcceptTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineTradeRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineTradeRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeclineTradeRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyLeave()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyLeave
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPartyLeave", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyRequest()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CPartyRequest
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendPartyRequest", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAcceptParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CAcceptParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendAcceptParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDeclineParty()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CDeclineParty
    SendData buffer.ToArray()
    Set buffer = Nothing

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendDeclineParty", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SendChatOption(Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong Index
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    InEvent = False
End Sub

Public Sub SendChatContinue()
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong 0
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    InEvent = False
End Sub

Public Sub EnviarMapaCubos(ByVal TileN As Long, ByVal TileX As Long, ByVal TileY As Long, ByVal Capa As Long, ByVal Tipo As Long, ByVal Mapa As Long, ByVal Mapax As Long, ByVal Mapay As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal Data4 As String, ByVal HP As Long, ByVal Animacion As Long, ByVal Banco As Long, ByVal Evento As Long, ByVal BancoLlave As Long, ByVal Script As Long, ByVal Timer As Long, ByVal SFX1 As Long, ByVal SFX2 As Long, ByVal Mensaje As String, ByVal Objeto As Long) 'Una supersintetizacion de SendData para incrementar rendimiento de memoria virtual
Dim packet As String
Dim buffer As clsBuffer
'FUNCION PARA CUBOS 32x32
procesandocubo = True
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    CanMoveNow = False 'Bloquea PJ
        
    buffer.WriteLong CEnviarMapaCubos
    buffer.WriteLong TileX
    buffer.WriteLong TileY
    buffer.WriteLong TileN
    buffer.WriteLong Capa
    buffer.WriteLong Mapax
    buffer.WriteLong Mapay
    buffer.WriteLong Mapa
    buffer.WriteLong Tipo
    buffer.WriteLong Data1 'Si es warp
    buffer.WriteLong Data2
    buffer.WriteLong Data3
    buffer.WriteString Data4
    buffer.WriteLong HP
    buffer.WriteLong Animacion
    buffer.WriteLong Banco
    buffer.WriteLong Evento
    buffer.WriteLong BancoLlave
    buffer.WriteLong Script
    buffer.WriteLong Timer
    buffer.WriteLong SFX1
    buffer.WriteLong SFX2
    buffer.WriteString Mensaje
    buffer.WriteLong Objeto
 
    SendData buffer.ToArray() 'Se envia con exito
    Set buffer = Nothing
    
    ClearTempTile
    initAutotiles
    
     Call ActualizarMapaCubos
     
procesandocubo = False
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EnviarMapaCubos64(ByVal TileN As Long, ByVal TileX As Long, ByVal TileY As Long, ByVal Capa As Long, ByVal Tipo As Long, ByVal Mapa As Long, ByVal Mapax As Long, ByVal Mapay As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal Data4 As String, ByVal HP As Long, ByVal Animacion As Long, ByVal Banco As Long, ByVal Evento As Long, ByVal BancoLlave As Long, ByVal Script As Long, ByVal Timer As Long, ByVal SFX1 As Long, ByVal SFX2 As Long, ByVal Mensaje1 As String, ByVal Objeto As Long, ByVal TileX2 As Long, ByVal TileY2 As Long, ByVal Capa2 As Long, ByVal Tipo2 As Long, ByVal Mapax2 As Long, ByVal Mapay2 As Long, ByVal Data12 As Long, ByVal Data22 As Long, ByVal Data32 As Long, ByVal Data42 As String, ByVal HP2 As Long, ByVal Animacion2 As Long, ByVal Banco2 As Long, ByVal Evento2 As Long, ByVal BancoLlave2 As Long, ByVal Script2 As Long, ByVal Timer2 As Long, ByVal SFX01 As Long, ByVal SFX02 As Long, ByVal Mensaje2 As String)
Dim packet As String
Dim buffer As clsBuffer
'FUNCION PARA CUBOS 32x64
procesandocubo = True

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    CanMoveNow = False 'Bloquea PJ
    
    buffer.WriteLong CEnviarMapaCubos64
    buffer.WriteLong TileX
    buffer.WriteLong TileY
    buffer.WriteLong TileN
    buffer.WriteLong Capa
    buffer.WriteLong Mapax
    buffer.WriteLong Mapay
    buffer.WriteLong Mapa
    buffer.WriteLong Tipo
    buffer.WriteLong Data1 'Si es warp
    buffer.WriteLong Data2
    buffer.WriteLong Data3
    buffer.WriteString Data4
    buffer.WriteLong HP
    buffer.WriteLong Animacion
    buffer.WriteLong Banco
    buffer.WriteLong Evento
    buffer.WriteLong BancoLlave
    buffer.WriteLong Script
    buffer.WriteLong Timer
    buffer.WriteLong SFX1
    buffer.WriteLong SFX2
    buffer.WriteString Mensaje1
    buffer.WriteLong Objeto
    
    buffer.WriteLong TileX2
    buffer.WriteLong TileY2
    buffer.WriteLong Capa2
    buffer.WriteLong Mapax2
    buffer.WriteLong Mapay2
    buffer.WriteLong Tipo2
    buffer.WriteLong Data12 'Si es warp
    buffer.WriteLong Data22
    buffer.WriteLong Data32
    buffer.WriteString Data42
    buffer.WriteLong HP2
    buffer.WriteLong Animacion2
    buffer.WriteLong Banco2
    buffer.WriteLong Evento2
    buffer.WriteLong BancoLlave2
    buffer.WriteLong Script2
    buffer.WriteLong Timer2
    buffer.WriteLong SFX01
    buffer.WriteLong SFX02
    buffer.WriteString Mensaje2



 
    SendData buffer.ToArray() 'Se envia con exito
    Set buffer = Nothing
       
       
    ClearTempTile
    initAutotiles
       
       
    Call ActualizarMapaCubos
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMap", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Public Sub RedibujarMapaCubos()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CMapaRespawnCubos
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendMapRespawn", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Golpe(ByVal Golpe As Long) 'Solicitud de inflingir damage, medio emo
  Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteLong CGolpe
    buffer.WriteLong Golpe
    SendData buffer.ToArray
    Set buffer = Nothing
    
End Sub

Public Sub EnviarVisibilidad() 'Restaura Jugador Visible
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEnviarVisibilidad
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendVisibility", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RestaurarBuff(ByVal StatNum As Byte, ByVal BuffVal As Byte)
Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    Set buffer = New clsBuffer
    buffer.WriteLong CEstadosBuff
    buffer.WriteByte StatNum
    buffer.WriteByte BuffVal
    SendData buffer.ToArray()
    Set buffer = Nothing
    
Select Case StatNum

Case 1

If BuffVal = 1 Then FBN = False
If BuffVal = 0 Then FBP = False

Case 2

If BuffVal = 1 Then DBN = False
If BuffVal = 0 Then DBP = False


Case 3

If BuffVal = 1 Then IBN = False
If BuffVal = 0 Then IBP = False


Case 4

If BuffVal = 1 Then ABN = False
If BuffVal = 0 Then ABP = False

Case 5

If BuffVal = 1 Then VBN = False
If BuffVal = 0 Then VBP = False



End Select

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendTrainStat", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Sub Restaurar_Sprite() 'EaSee 0.5
  Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CRestaurarSprite
    
    SendData buffer.ToArray()
    Set buffer = Nothing
    SpriteH = True
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SendSetSprite", "modClientTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Function EstaCorriendo(ByVal NombreDelProceso As String) As Boolean
Const MAX_PATH As Long = 260
Dim lProcesses() As Long, lModules() As Long, n As Long, lRet As Long, hProcess As Long
Dim sName As String
NombreDelProceso = UCase$(NombreDelProceso)
ReDim lProcesses(1023) As Long
If EnumProcesses(lProcesses(0), 1024 * 4, lRet) Then
For n = 0 To (lRet \ 4) - 1
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcesses(n))
If hProcess Then
ReDim lModules(1023)
If EnumProcessModules(hProcess, lModules(0), 1024 * 4, lRet) Then
sName = String$(MAX_PATH, vbNullChar)
GetModuleBaseName hProcess, lModules(0), sName, MAX_PATH
sName = Left$(sName, InStr(sName, vbNullChar) - 1)
If Len(sName) = Len(NombreDelProceso) Then
If NombreDelProceso = UCase$(sName) Then EstaCorriendo = True: Exit Function
End If
End If
End If
CloseHandle hProcess
Next n
End If
End Function

