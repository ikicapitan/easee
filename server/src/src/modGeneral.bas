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
        If Len(tempStr) > 0 Then .chkFriendSystem.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "DropOnDeath")
        If Len(tempStr) > 0 Then .chkDropInvItems.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "FullScreen")
        If Len(tempStr) > 0 Then .chkFS.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "Projectiles")
        If Len(tempStr) > 0 Then .chkProj.Value = Val(tempStr)
        tempStr = GetVar(Path, "OPTIONS", "OriginalGUIBars")
        If Len(tempStr) > 0 Then .chkGUIBars.Value = Val(tempStr)
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
    Call SetStatus("Iniciando Variables del Jugador...")

    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
        Load frmServer.Socket(I)
    Next

    ' Serves as a constructor
    Call ClearGameData
    Call LoadGameData
    Call SetStatus("Insertando Objetos en Mapas...")
    Call SpawnAllMapsItems
    Call SetStatus("Insertando NPCs en Mapas...")
    Call SpawnAllMapNpcs
    Call SetStatus("Insertando Eventos Globales...")
    Call SpawnAllMapGlobalEvents
    Call SetStatus("Creando Cache del Mapa...")
    Call CreateFullMapCache
    
    Call SetStatus("Cargando Bandeja del Sistema...")
    Call LoadSystemTray
    frmServer.tmrGetTime.Enabled = True
    Call SetStatus("Configurando Reloj")
    
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
        Call SetStatus("Bienvenido a Easee Engine. Servidor cargado en " & Output & " segundos.")
    Else
        Call SetStatus("Bienvenido a Easee Engine. Servidor cargado en " & time2 - time1 & " milisegundos.")
    End If
    
    ' reset shutdown value
    isShuttingDown = False
    
    ' Starts the server loop
    ServerLoop
End Sub

Public Sub DestroyServer()
    Dim I As Long
    ServerOnline = False
    Call SetStatus("Vaciando Bandeja del Sistema...")
    Call DestroySystemTray
    Call SetStatus("Guardando Datos de Jugadores Online...")
    Call SaveAllPlayersOnline
    Call ClearGameData
    Call SetStatus("Vaciando Sockets...")

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
    Call SetStatus("Limpiando Tiles Temporales...")
    Call ClearTempTiles
    Call SetStatus("Limpiando Mapas...")
    Call ClearMaps
    Call SetStatus("Limpiando Objetos en Mapas...")
    Call ClearMapItems
    Call SetStatus("Limpiando NPCs en Mapas...")
    Call ClearMapNpcs
    Call SetStatus("Limpiando NPCs...")
    Call ClearNpcs
    Call SetStatus("Limpiando Recursos...")
    Call ClearResources
    Call SetStatus("Limpiando Objetos...")
    Call ClearItems
    Call SetStatus("Limpiando Tiendas...")
    Call ClearShops
    Call SetStatus("Limpiando Hechizos...")
    Call ClearSpells
    Call SetStatus("Limpiando Animaciones...")
    Call ClearAnimations
    Call SetStatus("Limpiando Grupos...")
    Call ClearGuilds
    Call SetStatus("Limpiando Misiones...")
    Call ClearQuests
    Call SetStatus("Limpiando Combos...")
    Call ClearCombos
End Sub

Private Sub LoadGameData()
    Call SetStatus("Cargando Clases...")
    Call LoadClasses
    Call SetStatus("Cargando Mapas...")
    Call LoadMaps
    Call SetStatus("Cargando Objetos...")
    Call LoadItems
    Call SetStatus("Cargando NPCs...")
    Call LoadNpcs
    Call SetStatus("Cargando Recursos...")
    Call LoadResources
    Call SetStatus("Cargando Tiendas...")
    Call LoadShops
    Call SetStatus("Cargando Hechizos...")
    Call LoadSpells
    Call SetStatus("Cargando Animaciones...")
    Call LoadAnimations
    Call SetStatus("Cargando Switchs...")
    Call LoadSwitches
    Call SetStatus("Cargando Variables...")
    Call LoadVariables
    Call SetStatus("Cargando Misiones...")
    Call LoadQuests
    Call SetStatus("Cargando McCombos...")
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




