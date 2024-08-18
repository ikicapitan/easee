Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
Dim filename As String
Dim VTemp() As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    filename = App.Path & "\Data Files\config.ini"
    
    'VTemp = Split(frmMenu.ConfigCo2, "×")
    'Call PutVar(filename, "Resolucion", "FPS", frmMenu.ConfigCo1)
    'Call PutVar(filename, "Resolucion", "MODE", frmMenu.ConfigCo3)
    Call PutVar(filename, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(filename, "Options", "IP", Options.IP)
    Call PutVar(filename, "Options", "Port", str(Options.Port))
    Call PutVar(filename, "Options", "MenuMusic", Trim$(Options.MenuMusic))
    Call PutVar(filename, "Options", "IntroMusic", Trim$(Options.IntroMusic))
    Call PutVar(filename, "Options", "Music", str(Options.Music))
    Call PutVar(filename, "Options", "Sound", str(Options.sound))
    Call PutVar(filename, "Options", "Debug", str(Options.Debug))
    Call PutVar(filename, "Options", "Levels", str(Options.Lvls))
    Call PutVar(filename, "Options", "Buttons", str(Options.Buttons))
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
Dim filename As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    filename = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(filename, True) Then
        Options.Game_Name = "Easee Engine"
        Options.Password = vbNullString
        Options.savePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.IntroMusic = vbNullString
        Options.MenuMusic = vbNullString
        Options.Music = 1
        Options.sound = 1
        Options.Debug = 0
        Options.Lvls = 0
        Options.MiniMap = 0
        Options.Buttons = 0
        Options.Resol_Ancho = "0"
        Options.Resol_Alto = "0"
        SaveOptions
    Else
        Options.Game_Name = GetVar(filename, "Options", "Game_Name")
        If GetVar(App.Path & "\data files\config.ini", "Options", "SaveAccount") = 1 Then
        If GetVar(filename, "USER", "Username") = "" Then
        Else
        frmMenu.chkPass.Value = 1
        End If
        If GetVar(filename, "USER", "Username") = "" Then
        Else
        frmMenu.txtLUser = GetVar(filename, "USER", "Username")
        End If
        Else
        Call PutVar(filename, "Options", "Username", "")
        End If
        Options.savePass = Val(GetVar(filename, "Options", "SaveAccount"))
        Options.Port = Val(GetVar(filename, "Options", "Port"))
        Options.IntroMusic = GetVar(filename, "Options", "IntroMusic")
        Options.MenuMusic = GetVar(filename, "Options", "MenuMusic")
        Options.Music = GetVar(filename, "Options", "Music")
        Options.sound = GetVar(filename, "Options", "Sound")
        Options.Debug = GetVar(filename, "Options", "Debug")
        Options.Lvls = GetVar(filename, "Options", "Levels")
        Options.Buttons = GetVar(filename, "Options", "Buttons")
        Options.IP = GetVar(filename, "Options", "IP")
    
        End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim x As Long
Dim y As Long, I As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open filename For Binary As #f
    Put #f, , Map.name
    Put #f, , Map.Music
    Put #f, , Map.BGS
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    
    Put #f, , Map.Weather
    Put #f, , Map.WeatherIntensity
    
    Put #f, , Map.Fog
    Put #f, , Map.FogSpeed
    Put #f, , Map.FogOpacity
    
    Put #f, , Map.Red
    Put #f, , Map.Green
    Put #f, , Map.Blue
    Put #f, , Map.alpha
    
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Put #f, , Map.Tile(x, y)
        Next

        DoEvents
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map.NPC(x)
        Put #f, , Map.NpcSpawnType(x)
    Next
    
    Put #f, , Map.DropItemsOnDeath

    Close #f
    
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim filename As String
Dim f As Long
Dim x As Long
Dim y As Long, I As Long, Z As Long, w As Long, p As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    filename = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open filename For Binary As #f
    Get #f, , Map.name
    Get #f, , Map.Music
    Get #f, , Map.BGS
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    
    Get #f, , Map.Weather
    Get #f, , Map.WeatherIntensity
        
    Get #f, , Map.Fog
    Get #f, , Map.FogSpeed
    Get #f, , Map.FogOpacity
        
    Get #f, , Map.Red
    Get #f, , Map.Green
    Get #f, , Map.Blue
    Get #f, , Map.alpha
    
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY

    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Get #f, , Map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , Map.NPC(x)
        Get #f, , Map.NpcSpawnType(x)
    Next
    
    Get #f, , Map.DropItemsOnDeath

    Close #f
    ClearTempTile
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    I = 1
    NumTileSets = 1
    
    ReDim Tex_Tileset(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "tilesets\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Tileset(NumTileSets)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Tileset(NumTileSets).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "tilesets\" & I & GFX_EXT, "G", True)
        Tex_Tileset(NumTileSets).Texture = NumTextures
        NumTileSets = NumTileSets + 1
        I = I + 1
            'SEasee.ProC(CodeSv,
    ',"G",True)
    Wend
    
    NumTileSets = NumTileSets - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumCharacters = 1
    
    ReDim Tex_Character(1)
    

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "characters\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Character(NumCharacters)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Character(NumCharacters).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "characters\" & I & GFX_EXT, "G", False)
        Tex_Character(NumCharacters).Texture = NumTextures
        NumCharacters = NumCharacters + 1
        I = I + 1
        
        'SEasee.ProC(CodeSv,
    ',"G",True)
    Wend
    
    NumCharacters = NumCharacters - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumPaperdolls = 1
    
    ReDim Tex_Paperdoll(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "paperdolls\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Paperdoll(NumPaperdolls)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Paperdoll(NumPaperdolls).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "paperdolls\" & I & GFX_EXT, "G", False)
        Tex_Paperdoll(NumPaperdolls).Texture = NumTextures
        NumPaperdolls = NumPaperdolls + 1
        I = I + 1
    Wend
    
    NumPaperdolls = NumPaperdolls - 1
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumAnimations = 1
    
    ReDim Tex_Animation(1)
    ReDim AnimationTimer(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "animations\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Animation(NumAnimations)
        ReDim Preserve AnimationTimer(NumAnimations)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Animation(NumAnimations).Texture = NumTextures
        Tex_Animation(NumAnimations).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "animations\" & I & GFX_EXT, "G", False)
        NumAnimations = NumAnimations + 1
        I = I + 1
    Wend
    
    NumAnimations = NumAnimations - 1

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    numitems = 1
    
    ReDim Tex_Item(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "items\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Item(numitems)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item(numitems).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "items\" & I & GFX_EXT, "G", False)
        Tex_Item(numitems).Texture = NumTextures
        numitems = numitems + 1
        I = I + 1
    Wend
    
    numitems = numitems - 1

    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumResources = 1
    
    ReDim Tex_Resource(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "resources\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Resource(NumResources)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Resource(NumResources).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "resources\" & I & GFX_EXT, "G", False)
        Tex_Resource(NumResources).Texture = NumTextures
        NumResources = NumResources + 1
        I = I + 1
    Wend
    
    NumResources = NumResources - 1

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpellIcons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumSpellIcons = 1
    
    ReDim Tex_SpellIcon(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "spellicons\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_SpellIcon(NumSpellIcons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_SpellIcon(NumSpellIcons).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "spellicons\" & I & GFX_EXT, "G", False)
        Tex_SpellIcon(NumSpellIcons).Texture = NumTextures
        NumSpellIcons = NumSpellIcons + 1
        I = I + 1
    Wend
    
    NumSpellIcons = NumSpellIcons - 1

    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckSpellIcons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumFaces = 1
    
    ReDim Tex_Face(1)

    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "Faces\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Face(NumFaces)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Face(NumFaces).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "faces\" & I & GFX_EXT, "G", False)
        Tex_Face(NumFaces).Texture = NumTextures
        NumFaces = NumFaces + 1
        I = I + 1
    Wend
    
    NumFaces = NumFaces - 1

    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFogs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumFogs = 1
    
    ReDim Tex_Fog(1)
    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "fogs\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Fog(NumFogs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Fog(NumFogs).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "fogs\" & I & GFX_EXT, "G", False)
        Tex_Fog(NumFogs).Texture = NumTextures
        NumFogs = NumFogs + 1
        I = I + 1
    Wend
    
    NumFogs = NumFogs - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckFogs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckGUIs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumGUIs = 1
    
    ReDim Tex_GUI(1)
        While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_GUI(NumGUIs)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_GUI(NumGUIs).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\" & I & GFX_EXT, "G", False)
        Tex_GUI(NumGUIs).Texture = NumTextures
        NumGUIs = NumGUIs + 1
        I = I + 1
    Wend
    
    NumGUIs = NumGUIs - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckGUIs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumButtons = 1
    
    ReDim Tex_Buttons(1)
    While FileExist(SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\buttons\" & I & GFX_EXT, "G", True), True)
        ReDim Preserve Tex_Buttons(NumButtons)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons(NumButtons).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\buttons\" & I & GFX_EXT, "G", False)
        Tex_Buttons(NumButtons).Texture = NumTextures
        NumButtons = NumButtons + 1
        I = I + 1
    Wend
    
    NumButtons = NumButtons - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckButtons", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons_c()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumButtons_c = 1
    
    ReDim Tex_Buttons_c(1)
    While FileExist(SEasee.ProC(CodeSv, GFX_PATH & "gui\buttons\" & I & "_c" & GFX_EXT, "G", True))
        ReDim Preserve Tex_Buttons_c(NumButtons_c)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons_c(NumButtons_c).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\buttons\" & I & "_c" & GFX_EXT, "G", False)
        Tex_Buttons_c(NumButtons_c).Texture = NumTextures
        NumButtons_c = NumButtons_c + 1
        I = I + 1
                            'SEasee.ProC(CodeSv,
    ',"G",True)
    Wend
    
    NumButtons_c = NumButtons_c - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckButtons_c", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckItems_S()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumItems_S = 1
    
    ReDim Tex_Item_S(1)
    
    While FileExist(SEasee.ProC(CodeSv, GFX_PATH & "items\" & I & "_s" & GFX_EXT, "G", True))
        ReDim Preserve Tex_Item_S(NumItems_S)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Item_S(NumItems_S).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "items\" & I & "_s" & GFX_EXT, "G", False)
        Tex_Item_S(NumItems_S).Texture = NumTextures
        NumItems_S = NumItems_S + 1
        I = I + 1
    Wend
    
    NumItems_S = NumItems_S - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckItems_S", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub CheckButtons_h()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    I = 1
    NumButtons_h = 1
    
    ReDim Tex_Buttons_h(1)
    While FileExist(SEasee.ProC(CodeSv, GFX_PATH & "gui\buttons\" & I & "_h" & GFX_EXT, "G", True))
        ReDim Preserve Tex_Buttons_h(NumButtons_h)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Buttons_h(NumButtons_h).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "gui\buttons\" & I & "_h" & GFX_EXT, "G", False)
        Tex_Buttons_h(NumButtons_h).Texture = NumTextures
        NumButtons_h = NumButtons_h + 1
        I = I + 1
    Wend
    
    NumButtons_h = NumButtons_h - 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckButtons_h", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).name = vbNullString
    Item(Index).Desc = vbNullString
    Item(Index).sound = "Ninguno."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearCombo(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Combo(Index)), LenB(Combo(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearCombo", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub InvHidden()
Dim buffer As clsBuffer
    
    If GUIWindow(GUI_INVENTORY).Visible = False Then
        Set buffer = New clsBuffer
        buffer.WriteLong CInvHidden
        SendData buffer.ToArray()
        Set buffer = Nothing
    End If
End Sub

Sub ClearCombos()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_COMBO
        Call ClearCombo(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearCombos", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Animation(Index).name = vbNullString
    Animation(Index).sound = "Ninguno."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_ANIMATIONS
        Call ClearAnimation(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNPC(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).name = vbNullString
    NPC(Index).sound = "Ninguno."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_NPCS
        Call ClearNPC(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).name = vbNullString
    Spell(Index).Desc = vbNullString
    Spell(Index).sound = "Ninguno."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).name = vbNullString
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    Resource(Index).sound = "Ninguno."
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_RESOURCES
        Call ClearResource(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    initAutotiles
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(I)
    Next

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********************
' ** Player functions **
' **********************
Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).name)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).name = name
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).EXP
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerNextLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).EXP = EXP
    If GetPlayerLevel(Index) = MAX_LEVELS And Player(Index).EXP > GetPlayerNextLevel(Index) Then
        Player(Index).EXP = GetPlayerNextLevel(Index)
        Exit Sub
    End If
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVisible(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVisible = Player(Index).Visible
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVisible(ByVal Index As Long, ByVal Visible As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Visible = Visible
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)

    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerCoins(ByVal Index As Long) As Long
Dim I As Long, Cnt As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Cnt = 0

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) > 0 Then
            If Item(GetPlayerInvItemNum(Index, I)).Type = ITEM_TYPE_CURRENCY Then
                Cnt = Cnt + GetPlayerInvItemValue(Index, I)
            End If
        End If
    Next
    
    GetPlayerCoins = Cnt
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerCoins", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_BYTE Then Value = MAX_BYTE
    Player(Index).Stat(Stat) = Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).POINTS
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).POINTS = POINTS
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Or Index <= 0 Then Exit Function
    GetPlayerMap = Player(Index).Map
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).x
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).x = x
    MiniMapPlayer(Index).x = x * 4
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).y
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).y = y
    MiniMapPlayer(Index).y = y * 4
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Dir = Dir
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    If invSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(invSlot).num
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal invSlot As Long, ByVal itemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).num = itemNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(invSlot).Value
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Sub
    PlayerInv(invSlot).Value = ItemValue
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    'If Options.Debug = 1 Then On Error GoTo ErrorHandler
    On Error GoTo error:
    
    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot)
    
    ' Error handler
    Exit Function
error:
'ErrorHandler:
'    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
'    Err.Clear
'    Exit Function
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot) = invNum
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectiles
Public Sub CheckProjectiles()
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    I = 1
    NumProjectiles = 1
    
    ReDim Tex_Projectile(1)
    
    While FileExist(SEasee.ProC(CodeSv, GFX_PATH & "projectiles\" & I & GFX_EXT, "G", True))
        ReDim Preserve Tex_Projectile(NumProjectiles)
        NumTextures = NumTextures + 1
        ReDim Preserve gTexture(NumTextures)
        Tex_Projectile(NumProjectiles).Filepath = SEasee.ProC(CodeSv, App.Path & GFX_PATH & "projectiles\" & I & GFX_EXT, "G", True)
        Tex_Projectile(NumProjectiles).Texture = NumTextures
        NumProjectiles = NumProjectiles + 1
        I = I + 1
    Wend
    
    NumProjectiles = NumProjectiles - 1
    
    If NumProjectiles = 0 Then Exit Sub
    
    For I = 1 To NumProjectiles
        LoadTexture Tex_Projectile(I)
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearProjectile(ByVal Index As Long, ByVal PlayerProjectile As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With Player(Index).ProjecTile(PlayerProjectile)
        .Direction = 0
        .Pic = 0
        .TravelTime = 0
        .x = 0
        .y = 0
        .Range = 0
        .Damage = 0
        .speed = 0
        .Municion = 0
    End With
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ClearProjectile", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

