Attribute VB_Name = "modGeneral"
Option Explicit
' Get system uptime in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Public Sub Main()
    Call InitServer
End Sub

Public Sub SetData()
Dim tempStr As String, Path As String
    With frmServer
        Call .UsersOnline_Start
        '.cboColor_Start.ListIndex = 15
        '.cboColor_End.ListIndex = 15
        '.cboColor_ActionMsg.ListIndex = 4
        '.cboColor_PlayerMsg.ListIndex = 2
        SetMainSkillData
        
        'Control Data
        Path = App.Path & "\data\options.ini"
        tempStr = GetVar(Path, "OPTIONS", "FriendSystem")
        
        If Len(tempStr) > 0 Then
        .chkFriendSystem.Value = Val(tempStr)
        If tempStr = 1 Then
        .Pickcheck(0).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
        End If
        End If
        
        
        tempStr = GetVar(Path, "OPTIONS", "DropOnDeath")
        If Len(tempStr) > 0 Then
        .chkDropInvItems.Value = Val(tempStr)
        If tempStr = 1 Then
        .Pickcheck(3).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
        End If
        End If
        
        
        tempStr = GetVar(Path, "OPTIONS", "FullScreen")
        If Len(tempStr) > 0 Then .chkFS.Value = Val(tempStr)
        
        
        tempStr = GetVar(Path, "OPTIONS", "Projectiles")
        If Len(tempStr) > 0 Then
        .chkProj.Value = Val(tempStr)
        If tempStr = 1 Then
        .Pickcheck(1).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
        End If
        End If
        
        tempStr = GetVar(Path, "OPTIONS", "OriginalGUIBars")
        If tempStr = 0 Then
        .chkGUIBars.Value = Val(tempStr)
        If Val(tempStr) = 1 Then
        .Pickcheck(4).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
        End If
        End If
        
        tempStr = GetVar(Path, "OPTIONS", "AnimacionAtaque")
        If tempStr = 0 Then
        .chk5frames.Value = Val(tempStr)
        If Val(tempStr) = 1 Then
        .Pickcheck(6).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
        End If
        End If
    End With
End Sub

Public Sub InitServer()
    Dim I As Long
    Dim F As Long
    Dim time1 As Long
    Dim time2 As Long
    Call InitMessages
    time1 = GetTickCount
    frmServer.Show
    Dim var1
    ' Initialize the random-number generator
    Randomize ', seed

    ' Check if the directory is there, if its not make it
    ChkDir App.Path & "\Data\", "accounts"
    ChkDir App.Path & "\Data\", "animations"
    ChkDir App.Path & "\Data\", "banks"
    ChkDir App.Path & "\Data\", "items"
    ChkDir App.Path & "\Data\", "logs"
    ChkDir App.Path & "\Data\", "maps"
    ChkDir App.Path & "\Data\", "npcs"
    ChkDir App.Path & "\Data\", "resources"
    ChkDir App.Path & "\Data\", "shops"
    ChkDir App.Path & "\Data\", "spells"
    ChkDir App.Path & "\Data\", "guilds"
    ChkDir App.Path & "\Data\", "quests"
    ChkDir App.Path & "\Data\", "skills"

    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' load options, set if they dont exist
    If Not FileExist(App.Path & "\data\options.ini", True) Then
        Options.Game_Name = "Easee Engine"
        Options.Port = 7001
        Options.MOTD = "Bienvenido a Easee Engine."
        Options.Website = "http://www.easee.es"
        Options.Buy_Cost = 5000
        Options.Buy_Lvl = 20
        Options.Buy_Item = 1
        Options.Join_Cost = 1000
        Options.Join_Lvl = 20
        Options.Join_Item = 1
        SaveOptions
    Else
        LoadOptions
    End If
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = Options.Port
    
    ' Init all the player sockets
    
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s1")
    Call SetStatus(var1)        '[LOGS]s1

    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
        Load frmServer.Socket(I)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s2")
    Call SetStatus(var1)        '[LOGS]s2
    Call SpawnAllMapsItems
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s3")
    Call SetStatus(var1)        '[LOGS]s3
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s4")
    Call SetStatus(var1)        '[LOGS]s4
    Call SpawnAllMapNpcs
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s5")
    Call SetStatus(var1)        '[LOGS]s5
    Call SpawnAllMapGlobalEvents
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s6")
    Call SetStatus(var1)        '[LOGS]s6
    Call CreateFullMapCache
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s7")
    Call SetStatus(var1)        '[LOGS]s7
    Call LoadSystemTray
    frmServer.tmrGetTime.Enabled = True
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s8")
    Call SetStatus(var1)        '[LOGS]s8
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s9")
    Call SetStatus(var1)        '[LOGS]s9
    Call Datos_Easee
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    ' Setup Guild ranks
    Call Set_Default_Guild_Ranks

    ' Start listening
    frmServer.Socket(0).Listen
    Call UpdateCaption
    time2 = GetTickCount
    
    Dim Time3 As Double
    Dim Output As String
    Time3 = time2 - time1
    If Time3 > 1000 Then
        Time3 = Time3 / 1000
        Output = FormatNumber(Time3, 3)
        Call SetStatus("Bienvenido a EaSee MMORPG. Servidor cargado en " & Output & " segundos.")
    Else
        Call SetStatus("Bienvenido a EaSee MMORPG. Servidor cargado en " & time2 - time1 & " milisegundos.")
    End If
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim I As Long
    ServerOnline = False
    Dim var1 As String
    
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s10")
    Call SetStatus(var1)        '[LOGS]s10
    Call DestroySystemTray
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s11")
    Call SetStatus(var1)        '[LOGS]s11
    Call SaveAllPlayersOnline
    Call ClearGameData
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s12")
    Call SetStatus(var1)        '[LOGS]s12

    For I = 1 To MAX_PLAYERS
        Unload frmServer.Socket(I)
    Next

    End
End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(Status)
    DoEvents
End Sub

Public Sub ClearGameData()
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s13")
    Call SetStatus(var1)        '[LOGS]s13
    Call ClearTempTiles
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s14")
    Call SetStatus(var1)        '[LOGS]s14
    Call ClearMaps
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s15")
    Call SetStatus(var1)        '[LOGS]s15
    Call ClearMapItems
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s16")
    Call SetStatus(var1)        '[LOGS]s16
    Call ClearMapNpcs
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s17")
    Call SetStatus(var1)        '[LOGS]s17
    Call ClearNpcs
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s18")
    Call SetStatus(var1)        '[LOGS]s18
    Call ClearResources
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s19")
    Call SetStatus(var1)        '[LOGS]s19
    Call ClearItems
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s20")
    Call SetStatus(var1)        '[LOGS]s20
    Call ClearShops
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s21")
    Call SetStatus(var1)        '[LOGS]s21
    Call ClearSpells
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s22")
    Call SetStatus(var1)        '[LOGS]s22
    Call ClearAnimations
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s23")
    Call SetStatus(var1)        '[LOGS]s23
    Call ClearGuilds
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s24")
    Call SetStatus(var1)        '[LOGS]s24
    Call ClearQuests
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s25")
    Call SetStatus(var1)        '[LOGS]s25
    Call ClearCombos
End Sub

Private Sub LoadGameData()
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s26")
    Call SetStatus(var1)        '[LOGS]s26
    Call LoadClasses
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s27")
    Call SetStatus(var1)        '[LOGS]s27
    Call LoadMaps
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s28")
    Call SetStatus(var1)        '[LOGS]s28
    Call LoadItems
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s29")
    Call SetStatus(var1)        '[LOGS]s29
    Call LoadNpcs
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s30")
    Call SetStatus(var1)        '[LOGS]s30
    Call LoadResources
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s31")
    Call SetStatus(var1)        '[LOGS]s31
    Call LoadShops
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s32")
    Call SetStatus(var1)        '[LOGS]s32
    Call LoadSpells
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s33")
    Call SetStatus(var1)        '[LOGS]s33
    Call LoadAnimations
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s34")
    Call SetStatus(var1)        '[LOGS]s34
    Call LoadSwitches
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s35")
    Call SetStatus(var1)        '[LOGS]s35
    Call LoadVariables
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s36")
    Call SetStatus(var1)        '[LOGS]s36
    Call LoadQuests
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "s37")
    Call SetStatus(var1)        '[LOGS]s37
    Call LoadCombos
End Sub

Public Sub TextAdd(Msg As String)
    NumLines = NumLines + 1

    If NumLines >= MAX_LINES Then
        frmServer.txtText.Text = vbNullString
        NumLines = 0
    End If

    frmServer.txtText.Text = frmServer.txtText.Text & vbNewLine & Msg
    frmServer.txtText.SelStart = Len(frmServer.txtText.Text)
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean

    If (sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57) Then
        isNameLegal = True
    End If

End Function



