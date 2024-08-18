Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public mouseClicked As Boolean
Public mouseState As GUIType

Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If BMovimiento = "1" Then Exit Sub
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    If Not chatOn Then
        If Not GUIWindow(GUI_QUESTDIALOGUE).Visible Then 'fix quest and space key mix up
            If GetKeyState(vbKeySpace) < 0 Then
                CheckMapGetItem
            End If
        End If
        
         If GetKeyState(vbKeyZ) < 0 Then 'Atacar Cubo
        Call CheckAttackCubo 'EaSee 0.9
        Else
            
        End If
    
        'Move Up
        If GetKeyState(vbKeyUp) < 0 Or GetKeyState(vbKeyW) < 0 Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirUp = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Or GetKeyState(vbKeyD) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
            Exit Sub
        Else
            DirRight = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Or GetKeyState(vbKeyS) < 0 Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
            Exit Sub
        Else
            DirDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Or GetKeyState(vbKeyA) < 0 Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
            Exit Sub
        Else
            DirLeft = False
        End If
        
        
        If GetKeyState(vbKeyF12) < 0 Then
         ScreenshotMap 'EaSee 0.6
        End If
        
        If GetKeyState(vbKeyX) < 0 Then 'Tecla de insercion de Cubo Tecla EaSee
    'Cubo1
    If procesandocubo = False Then 'Si no se esta insertando un cubo actualmente (sobrecarga)
    Dim Objeto As Integer
    Dim x As Long
    Dim y As Long
    Dim TileX As Long
    Dim TileY As Long
    Dim tilenum As Long
    Dim CuboSupTipo, CuboInfTipo As Byte
    Dim Mapa As Integer
    Dim dato, Dato2, Dato3 As Byte
    Dim letrero As Boolean
    Dim Data4, Mensaje As String 'Lo usaremos tambien para el letrero
    Dim HP, Banco, Animacion, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Dropeo As Long
    'Cubo2
    Dim x2, y2, TileX2, TileY2, HP2, Banco2, Animacion2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02 As Long
    Dim Data42 As String
    Dim Dato02, Dato22, Dato32 As Byte
    
    
    letrero = False 'Flag
    Objeto = GetPlayerEquipment(MyIndex, Weapon) 'Chequea equipamiento
    
    
    If Objeto > 0 Then 'Chequea si hay algo equipado
    If Item(Objeto).Type = ITEM_TYPE_CUBO Then 'Si hay un cubo equipado

    'Datos Base
    
    x = GetPlayerX(MyIndex) 'Toma coordenadas del jugador
    y = GetPlayerY(MyIndex)
    CuboSupTipo = Item(Objeto).CuboSupTipo 'Atributo de Cubo
    CuboInfTipo = Item(Objeto).CuboInfTipo
    Mapa = GetPlayerMap(MyIndex)
    HP = Item(Objeto).CuboDureza
    'Los que faltan podrian ser incluidos en versiones futuras o pueden ser editados por ustedes
    Animacion = Item(Objeto).CuboAnimacion 'Para uso futuro
    Evento = 0
    Banco = 0
    BancoLlave = 0
    Script = 0
    Timer = 0
    SFX1 = Item(Objeto).CuboSFX1
    SFX2 = Item(Objeto).CuboSFX2
    SFX01 = Item(Objeto).CuboSFX1
    SFX02 = Item(Objeto).CuboSFX2
    Mensaje = ""
    Dropeo = Item(Objeto).CuboObjeto
    'Hasta aqui uso futuro
    
    
Select Case GetPlayerDir(MyIndex) 'Anulamos el proceso si hay limites de pantalla

Case DIR_UP

If GetPlayerY(MyIndex) = 0 Then Exit Sub

Case DIR_DOWN

If GetPlayerY(MyIndex) = Map.MaxY Then Exit Sub

Case DIR_LEFT

If GetPlayerX(MyIndex) = 0 Then Exit Sub

Case DIR_RIGHT

If GetPlayerX(MyIndex) = Map.MaxX Then Exit Sub

End Select

    
    
    
    'Direccion
    
    Select Case GetPlayerDir(MyIndex) 'Toma la direccion en que mira el personaje para insertar
        Case DIR_UP
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y - 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y - 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y - 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y - 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y - 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y - 1).Type, Mapa, x, y - 1, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)     'SendMap optimizado rendimiento
        End If
        
        
        
        
        ElseIf Item(Objeto).Cubo64 = True Then 'Funcion direccion arriba para Cubo de 32x64
        
        
        
        Map.Tile(x, y - 2).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y - 2).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y - 2).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y - 2).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y - 2).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y - 2).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y - 2).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y - 2).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
        Select Case CuboInfTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y - 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y - 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y - 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y - 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato2 = Item(Objeto).CuboMapa
        Dato22 = Item(Objeto).CuboMapaX
        Dato32 = Item(Objeto).CuboMapaY
        Data42 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y - 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato2 = Item(Objeto).CuboGolpe
        Dato22 = 0
        Dato32 = 0
        Data42 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 And CuboInfTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
         Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y - 2).Type, Mapa, x, y - 2, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x, y - 1).Type, x, y - 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje)
        End If
        
        
        End If
        
        
        
        Case DIR_DOWN 'Jugador mirando hacia abajo
                    
                    
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y + 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y + 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y + 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y + 1).Type, Mapa, x, y + 1, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)     'SendMap optimizado rendimiento
        End If
        
        
        
        
        ElseIf Item(Objeto).Cubo64 = True Then 'Funcion direccion arriba para Cubo de 32x64
        
        
        
        Map.Tile(x, y + 2).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y + 2).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y + 2).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y + 2).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y + 2).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y + 2).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y + 2).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y + 2).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
        Select Case CuboInfTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x, y + 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x, y + 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x, y + 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato2 = Item(Objeto).CuboMapa
        Dato22 = Item(Objeto).CuboMapaX
        Dato32 = Item(Objeto).CuboMapaY
        Data42 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato2 = Item(Objeto).CuboGolpe
        Dato22 = 0
        Dato32 = 0
        Data42 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 And CuboInfTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
         Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y + 2).Type, Mapa, x, y + 1, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x, y + 1).Type, x, y + 2, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje)
        End If
        
        
        End If
                        
                    
                    
                    
        Case DIR_LEFT 'Jugador direccion izquierda
        
        
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX  'Tile X
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x - 1, y).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x - 1, y).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x - 1, y).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x - 1, y).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x - 1, y).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y).Type, Mapa, x - 1, y, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)     'SendMap optimizado rendimiento
        End If
        
        
        
        
        ElseIf Item(Objeto).Cubo64 = True Then 'Funcion direccion arriba para Cubo de 32x64
        
        
        
        Map.Tile(x - 1, y + 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x - 1, y + 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x - 1, y + 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        Map.Tile(x - 1, y + 1).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x - 1, y + 1).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
        Select Case CuboInfTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x - 1, y).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x - 1, y).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x - 1, y).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x - 1, y).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato2 = Item(Objeto).CuboMapa
        Dato22 = Item(Objeto).CuboMapaX
        Dato32 = Item(Objeto).CuboMapaY
        Data42 = ""
        
        Case 4 'Trampa
        Map.Tile(x - 1, y).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato2 = Item(Objeto).CuboGolpe
        Dato22 = 0
        Dato32 = 0
        Data42 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 And CuboInfTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y + 1).Type, Mapa, x - 1, y, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y).Type, x - 1, y + 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje)
        End If
        
        
        End If
        
        
        
        
        
        Case DIR_RIGHT 'Personaje hacia la derecha
        
        
        
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX  'Tile X
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x + 1, y).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x + 1, y).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x + 1, y).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x + 1, y).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x + 1, y).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y).Type, Mapa, x + 1, y, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)     'SendMap optimizado rendimiento
        End If
        
        
        
        
        ElseIf Item(Objeto).Cubo64 = True Then 'Funcion direccion arriba para Cubo de 32x64
        
        
        
        Map.Tile(x + 1, y + 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x + 1, y + 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x + 1, y + 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        Map.Tile(x + 1, y + 1).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x + 1, y + 1).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
        Select Case CuboInfTipo
        
        Case 0 'Sin Atributo
        Map.Tile(x + 1, y).Type = TILE_TYPE_WALKABLE 'Traspasable
        
        Case 1 'Si es Bloqueo
        Map.Tile(x + 1, y).Type = TILE_TYPE_BLOCKED 'Dibuja Bloqueo
    
        Case 2 'Banco
        Map.Tile(x + 1, y).Type = TILE_TYPE_BANK 'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x + 1, y).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato2 = Item(Objeto).CuboMapa
        Dato22 = Item(Objeto).CuboMapaX
        Dato32 = Item(Objeto).CuboMapaY
        Data42 = ""
        
        Case 4 'Trampa
        Map.Tile(x + 1, y).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato2 = Item(Objeto).CuboGolpe
        Dato22 = 0
        Dato32 = 0
        Data42 = ""
        
        Case 5 'Mensaje
        frmMain.frmeditorletrero.Visible = True
        letrero = True
        CanMoveNow = False 'Bloquea PJ
        End Select
                
        
        If y > 1 And CuboSupTipo <> 5 And CuboInfTipo <> 5 Then 'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y + 1).Type, Mapa, x + 1, y, dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y).Type, x + 1, y + 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje)
        End If
        
        
        End If
        

           
            
            
            
            
    End Select

    
    End If
    End If
    End If
    End If
    End If
   
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim name As String
Dim I As Long
Dim n As Long
Dim Command() As String
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'Rodo
    If GetKeyState(vbKeyEscape) < 0 And Not chatOn Then
    If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then
    OpenGuiWindow 12
    Else
    
    End If
    frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False
    GUIWindow(GUI_GUILD).Visible = False
    End If
    
    If BTeclas = 1 Then Exit Sub
    chatText = MyText
    
    If GUIWindow(GUI_CURRENCY).Visible Then
        If (KeyAscii = vbKeyBack) Then
            If LenB(sDialogue) > 0 Then sDialogue = Mid$(sDialogue, 1, Len(sDialogue) - 1)
        End If
            
        ' And if neither, then add the character to the user's text buffer
        If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
            sDialogue = sDialogue & ChrW$(KeyAscii)
        End If
    End If
    
        
    
    'Controles Teclado de la GUI
         If GetKeyState(vbKeyI) < 0 And Not chatOn Then
         If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False Then
         OpenGuiWindow 3
         End If
         End If
    
         If GetKeyState(vbKeyL) < 0 And Not chatOn Then
         If GUIWindow(GUI_ACHIEVEMENTS).Visible = True Then
         GUIWindow(GUI_ACHIEVEMENTS).Visible = False
         Else
         If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False Then
         Call OrderAchievements(1)
         GUIWindow(GUI_ACHIEVEMENTS).Visible = True
         End If
         End If
         End If
    
    If GetKeyState(vbKeyE) < 0 And Not chatOn Then
    'vamos a chequear que no haya editores abiertos para habilitar la tecla (EaSee Engine 0.4)
    If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then
        If GetPlayerAccess(MyIndex) >= ADMIN_MONITOR Then
            ' Weza gonna rotate through editors
            If Scroll_Editor < 10 Then
                Scroll_Editor = Scroll_Editor + 1
            Else
                Scroll_Editor = 1
            End If
            
            Call ScrollEditor
        End If
    End If
    End If
    
    
    If GetKeyState(vbKeyO) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

         OpenGuiWindow 2
    End If
    End If
    
    If GetKeyState(vbKeyT) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            SendTradeRequest
        Else
            AddText "Objetivo inválido.", BrightRed
        End If
        End If
    End If
    If GetKeyState(vbKeyP) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

         OpenGuiWindow 6
    End If
    End If
    If GetKeyState(vbKeyG) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

         OpenGuiWindow 7
    End If
    End If
    If GetKeyState(vbKeyQ) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

         OpenGuiWindow 8
    End If
    End If
    If GetKeyState(vbKeyK) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

         OpenGuiWindow 9
    End If
    End If
    If GetKeyState(vbKeyF) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
            Call SendFollowPlayer(myTarget)
        End If
        End If
    End If
    If GetKeyState(vbKeyB) < 0 And Not chatOn Then
        If frmMain.picAdmin.Visible = False And frmEditor_Animation.Visible = False And frmEditor_Character.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Combos.Visible = False And frmEditor_Events.Visible = False And frmEditor_Item.Visible = False And frmEditor_Map.Visible = False And frmEditor_MapProperties.Visible = False And frmEditor_NPC.Visible = False And frmEditor_Quest.Visible = False And frmEditor_Resource.Visible = False And frmEditor_Shop.Visible = False And frmEditor_Spell.Visible = False And frmMain.frmeditorletrero.Visible = False And GUIWindow(GUI_GUILD).Visible = False Then

        Dim SendClick As Boolean
        If Not myTargetType = TARGET_TYPE_PLAYER Or myTarget = MyIndex Then
             OpenGuiWindow 10
        ElseIf myTarget > 0 And myTarget <> MyIndex Then
            'For adding friends
            SendClick = True
            For I = 1 To Player(MyIndex).Friends.count
                If Trim$(Player(MyIndex).Friends.NameOfFriend(I)) = GetPlayerName(myTarget) Then
                     OpenGuiWindow 10
                    SendClick = False
                    Exit For
                End If
            Next I
            
            If SendClick Then Call SendDblClickPos
        End If
        End If
    End If
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
        chatOn = Not chatOn
                
        'Guild Message
        If Left$(chatText, 1) = ";" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
        
            If Len(chatText) > 0 Then
                Call GuildMsg(chatText)
            End If
        
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If
        
        ' Broadcast message
        If Left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If Left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If Left$(chatText, 1) = "!" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(chatText)

                If Mid$(chatText, I, 1) <> Space(1) Then
                    name = name & Mid$(chatText, I, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, I, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) > 0 Then 'had - i)
                'I don't even know what this next line was meant for but it's useless
                'MyText = Mid$(chatText, i + 1, Len(chatText) - i)
                'MyText isn't used again other than to set it to nothing anyway.. people these days :p
                
                ' Send the message to the player
                Call PrivateMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case LCase(Command(0))
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /info, /who, /fps, /fpslock", HelpColor)
                
                Case "/guild"
                    If UBound(Command) < 1 Then
                        ' OpenGuiWindow 7
                        Call AddText("Comandos de la Guild:", HelpColor)
                        Call AddText("Crear Guild: /guild make (Nombre) (Tag)", HelpColor)
                        Call AddText("Para transferir el mando del clan /guild founder (name)", HelpColor)
                        Call AddText("Para invitar: /guild invite (name)", HelpColor)
                        Call AddText("Para irse del clan: /guild leave", HelpColor)
                        Call AddText("Abrir el panel del clan: /guild admin", HelpColor)
                        Call AddText("Expulsar a alguien del clan: /guild kick (name)", HelpColor)
                        Call AddText("Borrar clan: /guild disband yes", HelpColor)
                        Call AddText("Ver Clan: /guild view (online/all/offline)", HelpColor)
                        Call AddText("Para hablar con tu clan: ;Mensaje ", HelpColor)
                        GoTo Continue
                    End If
                
                Select Case LCase(Command(1))
                    Case "make"
                        If UBound(Command) = 3 Then
                            Call GuildMake(Command(2), Command(3))
                        Else
                            Call AddText("Comando Desconocido /guild make (name) (tag)", BrightRed)
                        End If
                    
                    Case "invite"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(2, Command(2))
                        Else
                            Call AddText("Comando Desconocido /guild invite (name)", BrightRed)
                        End If
                    
                    Case "leave"
                        Call GuildCommand(3, "")
                    
                    Case "admin"
                        Call GuildCommand(4, "")
                    
                    Case "view"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(5, Command(2))
                        Else
                            Call GuildCommand(5, "")
                        End If
                    
                    Case "accept"
                        Call GuildCommand(6, "")
                    
                    Case "decline"
                        Call GuildCommand(7, "")
                    
                    Case "founder"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(8, Command(2))
                        Else
                            Call AddText("Comando Desconocido /guild founder (name)", BrightRed)
                        End If
                    Case "kick"
                        If UBound(Command) = 2 Then
                            Call GuildCommand(9, Command(2))
                        Else
                            Call AddText("Comando Desconocido /guild kick (name)", BrightRed)
                        End If
                    Case "disband"
                        If UBound(Command) = 2 Then
                            If LCase(Command(2)) = LCase("yes") Then
                                Call GuildCommand(10, "")
                            Else
                                Call AddText("Tienes que escribir /guild disband yes (Para evitar lamentos!)", BrightRed)
                            End If
                        Else
                            Call AddText("Tienes que escribir /guild disband yes (Para evitar lamentos!)", BrightRed)
                    End If
                
                End Select
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If

                    Set buffer = New clsBuffer
                    buffer.WriteLong CPlayerInfoRequest
                    buffer.WriteString Command(1)
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set buffer = New clsBuffer
                    buffer.WriteLong CGetStats
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    'Walkthrough toggle
                Case "/walkthrough"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue
                    SendWalkthrough
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    BLoc = Not BLoc
                    ' Spawn item
                Case "/spawnitem"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    If UBound(Command) < 2 Then
                        AddText "Usage: /spawnitem (ItemNum) (ItemValue)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Trim$(Command(1))) Then
                        AddText "ItemNum must be numeric", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Trim$(Command(2))) Then
                        AddText "ItemValue must be numeric", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSpawnItem CLng(Command(1)), CLng(Command(2))
                    ' Bank
                Case "/bank"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue
                    
                    SendOpenBankCommand
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If

                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If

                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If

                    SendSetSprite CLng(Command(1)), GetPlayerName(MyIndex)
                    'visibility toggle
                    Case "/visible"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                    
                    SendVisibility
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If

                    SendBan Command(1)
                    ' Killing a player
                Case "/kill"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo Continue

                        If UBound(Command) < 1 Then
                            AddText "Usage: /kill (name)", AlertColor
                            GoTo Continue
                        End If

                    SendKillPlayer Command(1)
                    ' Level up player
                Case "/level"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                        If UBound(Command) < 1 Then
                            AddText "Usage: /level (name)", AlertColor
                            GoTo Continue
                        End If

                    SendRequestLevelUp Command(1)
                    ' // Developer Admin Commands //
                    ' Character Editor
                Case "/editchar"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    
                    frmEditor_Character.Visible = True
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditItem
                    ' Editing combo request
                Case "/editcombo"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditCombo
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue

                    SendRequestEditSpell
                Case "/editquest"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                    SendRequestEditQuest
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo Continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    SendCommand (MyText)
                    'AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
Continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    If Not chatOn Then Exit Sub
    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Public Sub HandleMouseMove(ByVal x As Long, ByVal y As Long, ByVal Button As Long)
Dim I As Long
On Error GoTo error:
    ' Set the global cursor position
    
    GlobalX = x
    GlobalY = y
    'RODO
    GlobalX_Map = (TileView.Left * PIC_X) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.Top * PIC_Y) + Camera.Top
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (x >= GUIWindow(I).x And x <= GUIWindow(I).x + GUIWindow(I).Width) And (y >= GUIWindow(I).y And y <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_CHAT, GUI_BARS ', GUI_MENU
                            ' Put nothing here and we can click through them!
                        Case GUI_INVENTORY, GUI_SPELLS, GUI_CHARACTER, GUI_PARTY, GUI_OPTIONS, GUI_GUILD, GUI_QUESTLOG, GUI_COMBAT
                            ' Moveable GUI if right clicked
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.Left + ((x + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((y + Camera.Top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, x, y)
        End If
    End If
    
    If mouseClicked Then
        If mouseState = GUI_INVENTORY Then
            GUIWindow(GUI_INVENTORY).x = GlobalX
            GUIWindow(GUI_INVENTORY).y = GlobalY
        ElseIf mouseState = GUI_SPELLS Then
            GUIWindow(GUI_SPELLS).x = GlobalX
            GUIWindow(GUI_SPELLS).y = GlobalY
        ElseIf mouseState = GUI_CHARACTER Then
            GUIWindow(GUI_CHARACTER).x = GlobalX
            GUIWindow(GUI_CHARACTER).y = GlobalY
        ElseIf mouseState = GUI_PARTY Then
            GUIWindow(GUI_PARTY).x = GlobalX
            GUIWindow(GUI_PARTY).y = GlobalY
        ElseIf mouseState = GUI_OPTIONS Then
            GUIWindow(GUI_OPTIONS).x = GlobalX
            GUIWindow(GUI_OPTIONS).y = GlobalY
        ElseIf mouseState = GUI_GUILD Then
            GUIWindow(GUI_GUILD).x = GlobalX
            GUIWindow(GUI_GUILD).y = GlobalY
        ElseIf mouseState = GUI_QUESTLOG Then
            GUIWindow(GUI_QUESTLOG).x = GlobalX
            GUIWindow(GUI_QUESTLOG).y = GlobalY
            frmMain.lstQuestLog.Left = (GUIWindow(GUI_QUESTLOG).x + (GUIWindow(GUI_QUESTLOG).Width / 2)) - (frmMain.lstQuestLog.Width / 2)
            frmMain.lstQuestLog.Top = GUIWindow(GUI_QUESTLOG).y + 10
        ElseIf mouseState = GUI_COMBAT Then
            GUIWindow(GUI_COMBAT).x = GlobalX
            GUIWindow(GUI_COMBAT).y = GlobalY
        ElseIf mouseState = GUI_FRIENDS Then
            GUIWindow(GUI_FRIENDS).x = GlobalX
            GUIWindow(GUI_FRIENDS).y = GlobalY
            frmMain.lstFriends.Left = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width / 2)) - (frmMain.lstFriends.Width / 2)
            frmMain.lstFriends.Top = GUIWindow(GUI_FRIENDS).y + 10
        End If
    End If

error:
End Sub
Public Sub HandleMouseDown(ByVal Button As Long)
Dim I As Long
'Msgbox Button
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_ACHIEVEMENTS
                             Achievement_MouseDown
                        Case GUI_CHAT
                            ' Put nothing here and we can click through the
                        Case GUI_MENUOPTIONS
                            OptionsMenu_MouseDown Button
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            mouseState = GUI_INVENTORY
                            'Msgbox Button
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            mouseState = GUI_SPELLS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_MENU
                            If Options.Buttons = 1 Then
                                Menu_MouseDown Button
                            End If
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            mouseState = GUI_CHARACTER
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_CURRENCY
                            Currency_MouseDown
                            Exit Sub
                        Case GUI_DIALOGUE
                            Dialogue_MouseDown
                            Exit Sub
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            mouseState = GUI_PARTY
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_OPTIONS
                            Options_MouseDown
                            mouseState = GUI_OPTIONS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_TRADE
                            Trade_MouseDown
                            Exit Sub
                        Case GUI_EVENTCHAT
                            Chat_MouseDown
                            Exit Sub
                        Case GUI_GUILD
                            Guild_MouseDown
                            mouseState = GUI_GUILD
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_QUESTLOG
                            QuestLog_MouseDown
                            mouseState = GUI_QUESTLOG
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseDown
                            Exit Sub
                        Case GUI_COMBAT
                            Combat_MouseDown
                            mouseState = GUI_COMBAT
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_FRIENDS
                            Buddies_MouseDown
                            mouseState = GUI_FRIENDS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_FRIENDREQUEST
                            FriendRequest_MouseDown
                            mouseState = GUI_FRIENDREQUEST
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_PLAYERINFO
                            PlayerInfo_MouseDown
                            mouseState = GUI_PLAYERINFO
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_BOOK
                            Book_MouseDown
                            mouseState = GUI_BOOK
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        Case GUI_BARS
                            Bars_MouseDown Button
                            mouseState = GUI_BARS
                            If Button = vbRightButton Then mouseClicked = True
                            Exit Sub
                        'QUICKCHANGE
                        'Case Else
                        '    Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If Not inChat Then
            ChatScroll_MouseDown
        End If
    End If
    
    If frmMenu.Visible = True Then
        MenuButton_MouseDown
    End If
    ' Handle events
    If InMapEditor Then
        Call MapEditorMouseDown(Button, GlobalX, GlobalY, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            Call PlayerSearch(CurX, CurY)
            'FindTarget
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If
    If frmEditor_Events.Visible Then frmEditor_Events.SetFocus
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim I As Long
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_CHAT
                            ' Put nothing here and we can click through the
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                            mouseClicked = False
                        Case GUI_SPELLS
                            Spells_MouseUp
                            mouseClicked = False
                        Case GUI_MENU
                            If Options.Buttons = 1 Then
                                Menu_MouseUp
                            End If
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                            mouseClicked = False
                        Case GUI_CURRENCY
                            Currency_MouseUp
                        Case GUI_DIALOGUE
                            Dialogue_MouseUp
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                            mouseClicked = False
                        Case GUI_OPTIONS
                            Options_MouseUp
                            mouseClicked = False
                        Case GUI_TRADE
                            Trade_MouseUp
                        Case GUI_EVENTCHAT
                            Chat_MouseUp
                        Case GUI_GUILD
                            Guild_MouseUp
                            mouseClicked = False
                        Case GUI_QUESTLOG
                            QuestLog_MouseUp
                            mouseClicked = False
                        Case GUI_QUESTDIALOGUE
                            QuestDialogue_MouseUp
                        Case GUI_COMBAT
                            Combat_MouseUp
                            mouseClicked = False
                        Case GUI_FRIENDS
                            Buddies_MouseUp
                            mouseClicked = False
                        Case GUI_FRIENDREQUEST
                            FriendRequest_MouseUp
                            mouseClicked = False
                        Case GUI_PLAYERINFO
                            PlayerInfo_MouseUp
                            mouseClicked = False
                        Case GUI_BARS
                            Bars_MouseUp
                            mouseClicked = False
                    End Select
                End If
            End If
        Next
    End If
    If frmMenu.Visible = True Then
    MenuButton_MouseUp
    End If
    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleSingleClick()
    If Not InMapEditor And Not hideGUI Then
        If (GlobalX >= GUIWindow(GUI_INVENTORY).x And GlobalX <= GUIWindow(GUI_INVENTORY).x + GUIWindow(GUI_INVENTORY).Width) And (GlobalY >= GUIWindow(GUI_INVENTORY).y And GlobalY <= GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_INVENTORY).Height) Then
            If GUIWindow(GUI_INVENTORY).Visible Then
                Inventory_SingleClick
            End If
        End If
    End If
End Sub

Public Sub HandleDoubleClick()
Dim I As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For I = 1 To GUI_Count - 1
            If (GlobalX >= GUIWindow(I).x And GlobalX <= GUIWindow(I).x + GUIWindow(I).Width) And (GlobalY >= GUIWindow(I).y And GlobalY <= GUIWindow(I).y + GUIWindow(I).Height) Then
                If GUIWindow(I).Visible Then
                    Select Case I
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                            Exit Sub
                        Case GUI_SHOP
                            Shop_DoubleClick
                            Exit Sub
                        Case GUI_BANK
                            Bank_DoubleClick
                            Exit Sub
                        Case GUI_TRADE
                            Trade_DoubleClick
                            Exit Sub
                        'Case Else
                        '    Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal Index As Long)
Dim buffer As clsBuffer
If frmMenu.Visible = True Then Exit Sub
    If Index = 1 Then
        GUIWindow(GUI_INVENTORY).Visible = Not GUIWindow(GUI_INVENTORY).Visible
        Call InvHidden
    Else
        GUIWindow(GUI_INVENTORY).Visible = False
        Call InvHidden
    End If
    
    If Index = 2 Then
        GUIWindow(GUI_SPELLS).Visible = Not GUIWindow(GUI_SPELLS).Visible
        ' Update the spells on the pic
        Set buffer = New clsBuffer
        buffer.WriteLong CSpells
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        GUIWindow(GUI_SPELLS).Visible = False
    End If
    
    If Index = 3 Then
        GUIWindow(GUI_CHARACTER).Visible = Not GUIWindow(GUI_CHARACTER).Visible
    Else
        GUIWindow(GUI_CHARACTER).Visible = False
    End If
    
    If Index = 4 Then
        GUIWindow(GUI_OPTIONS).Visible = Not GUIWindow(GUI_OPTIONS).Visible
    Else
        GUIWindow(GUI_OPTIONS).Visible = False
    End If
    
    If Index = 6 Then
        GUIWindow(GUI_PARTY).Visible = Not GUIWindow(GUI_PARTY).Visible
    Else
        GUIWindow(GUI_PARTY).Visible = False
    End If
    
    If Index = 7 Then
            If GUIWindow(GUI_GUILD).Visible = False Then
            'MAINKRA
            If Not Player(MyIndex).GuildName = vbNullString Then
            GUIWindow(GUI_GUILD).Visible = True
            Call GuildCommand(4, "")
            frmMain.ClanName.Visible = False
            frmMain.ClanTag.Visible = False
            frmMain.ClanBoton.Visible = False
            BTeclas = 0
            BMovimiento = 0
            Else
            GUIWindow(GUI_GUILD).Visible = True
            frmMain.ClanName.Visible = True
            frmMain.ClanTag.Visible = True
            frmMain.ClanBoton.Visible = True
            BTeclas = 1
            BMovimiento = 1
            End If
            Else
            GUIWindow(GUI_GUILD).Visible = False
            End If
    End If
    
    If Index = 8 Then
        GUIWindow(GUI_QUESTLOG).Visible = Not GUIWindow(GUI_QUESTLOG).Visible
        frmMain.lstQuestLog.Visible = Not frmMain.lstQuestLog.Visible
        UpdateQuestLog
    Else
        GUIWindow(GUI_QUESTLOG).Visible = False
        frmMain.lstQuestLog.Visible = False
    End If
    
    If Index = 9 Then
        GUIWindow(GUI_COMBAT).Visible = Not GUIWindow(GUI_COMBAT).Visible
    Else
        GUIWindow(GUI_COMBAT).Visible = False
    End If
    
    If Index = 10 Then
        GUIWindow(GUI_FRIENDS).Visible = Not GUIWindow(GUI_FRIENDS).Visible
        frmMain.lstFriends.Visible = Not frmMain.lstFriends.Visible
        UpdateFriendsList
    Else
        GUIWindow(GUI_FRIENDS).Visible = False
        frmMain.lstFriends.Visible = False
    End If
    
    If Index = 11 Then
        GUIWindow(GUI_MOUSE).Visible = True
        'RenderTexture Tex_MenuBg, GlobalX, GlobalY, GlobalX, GlobalY, 500, 500, 500, 500
         'RenderTexture Tex_MenuBg2, 0, 0, PosMenu, 0, 800, 600, 800, 600
    End If
    
    If Index = 12 Then
      If GUIWindow(GUI_MENUOPTIONS).Visible = False Then
       GUIWindow(GUI_MENUOPTIONS).Visible = True
      Else
       GUIWindow(GUI_MENUOPTIONS).Visible = False
    End If
   End If
End Sub

Public Sub Currency_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        CurrencyAcceptState = 2 ' clicked
    End If
    
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        CurrencyCloseState = 2 ' clicked
    End If
End Sub
Public Sub Currency_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long, buffer As clsBuffer
    Width = EngineGetTextWidth(Font_Default, "[Accept]")
    x = GUIWindow(GUI_CURRENCY).x + 155
    y = GUIWindow(GUI_CURRENCY).y + 96
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        If CurrencyAcceptState = 2 Then
            ' do stuffs
            If IsNumeric(sDialogue) Then
                Select Case CurrencyMenu
                    Case 1 ' drop item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        SendDropItem tmpCurrencyItem, Val(sDialogue)
                    Case 2 ' deposit item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        DepositItem tmpCurrencyItem, Val(sDialogue)
                    Case 3 ' withdraw item
                        If Val(sDialogue) > GetBankItemValue(tmpCurrencyItem) Then sDialogue = GetBankItemValue(tmpCurrencyItem)
                        WithdrawItem tmpCurrencyItem, Val(sDialogue)
                    Case 4 ' offer trade item
                        If Val(sDialogue) > GetPlayerInvItemValue(MyIndex, tmpCurrencyItem) Then sDialogue = GetPlayerInvItemValue(MyIndex, tmpCurrencyItem)
                        TradeItem tmpCurrencyItem, Val(sDialogue)
                End Select
            Else
                AddText "Ingresa Cantidad Valida.", BrightRed
                Exit Sub
            End If
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    Width = EngineGetTextWidth(Font_Default, "[Close]")
    x = GUIWindow(GUI_CURRENCY).x + 218
    y = GUIWindow(GUI_CURRENCY).y + 96
    ' check if we're on the button
    If (GlobalX >= x And GlobalX <= x + Buttons(12).Width) And (GlobalY >= y And GlobalY <= y + Buttons(12).Height) Then
        If CurrencyCloseState = 2 Then
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    CurrencyAcceptState = 0
    CurrencyCloseState = 0
    GUIWindow(GUI_CURRENCY).Visible = False
    inChat = False
    chatOn = False
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear
    sDialogue = vbNullString
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Dialogue_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 90
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(1) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(2) = 2 ' clicked
        End If
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        x = GUIWindow(GUI_DIALOGUE).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_DIALOGUE).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            Dialogue_ButtonState(3) = 2 ' clicked
        End If
    End If
End Sub

Public Sub Dialogue_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    If Dialogue_ButtonVisible(1) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Accept]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 90
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(1) = 2 Then
                Dialogue_Button_MouseDown (2)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(1) = 0
    End If
    If Dialogue_ButtonVisible(2) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Okay]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(2) = 2 Then
                Dialogue_Button_MouseDown (1)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(2) = 0
    End If
    If Dialogue_ButtonVisible(3) = True Then
        Width = EngineGetTextWidth(Font_Default, "[Close]")
        x = GUIWindow(GUI_CHAT).x + 10 + (155 - (Width / 2))
        y = GUIWindow(GUI_CHAT).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If Dialogue_ButtonState(3) = 2 Then
                Dialogue_Button_MouseDown (3)
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        Dialogue_ButtonState(3) = 0
    End If
End Sub

Public Sub MenuButton_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = 66 To 89
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).Visible = True Then
                Buttons(I).state = 2 ' clicked
            End If
        End If
    Next
End Sub

Public Sub MenuButton_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
Dim name As String
Dim Password As String
Dim PasswordAgain As String
Dim VTemp() As String
Dim dato As String
Dim ix As Integer, iy As Integer

    If FinIntro = False Then Exit Sub
    For I = 66 To 89
        x = Buttons(I).x
        y = Buttons(I).y
        Buttons(I).Visible = False
        ' Colision
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
           If Buttons(I).state = 2 Then ' Clicado
             'Exit Sub
             'RenderText Font_Default, GetVar(App.Path & "\data files\config.ini", "Resolucion", "FPS"), Buttons(74).x + 30, Buttons(74).y, White, 0, False
            'RenderText Font_Default, GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "X" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT"), Buttons(76).x + 20, Buttons(76).y, White, 0, False
            'RenderText Font_Default, GetVar(App.Path & "\data files\config.ini", "Resolucion", "MODE"), Buttons(78).x + 20, Buttons(78).y, White, 0, False
                Select Case I
                    Case 82
                    Call PutVar(App.Path & "\data files\config.ini", "Options", "Sound", "1")
                    Case 83
                    Call PutVar(App.Path & "\data files\config.ini", "Options", "Sound", "0")
                    Case 84
                    Call PutVar(App.Path & "\data files\config.ini", "Options", "Music", "1")
                    Case 85
                    Call PutVar(App.Path & "\data files\config.ini", "Options", "Music", "0")
                    Case 86
                    If GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolS") = "0" Then Exit Sub
                    Call PutVar(App.Path & "\data files\config.ini", "SONIDO", "VolS", Val(GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolS")) - Val(10))
                    Case 87
                    If 150 <= GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolS") Then Exit Sub
                    Call PutVar(App.Path & "\data files\config.ini", "SONIDO", "VolS", Val(GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolS")) + Val(10))
                    Case 88
                    If GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolM") <= "0" Then Exit Sub
                    Call PutVar(App.Path & "\data files\config.ini", "SONIDO", "VolM", Val(GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolM")) - Val(10))
                    Case 89
                    If 150 <= GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolM") Then Exit Sub
                    Call PutVar(App.Path & "\data files\config.ini", "SONIDO", "VolM", Val(GetVar(App.Path & "\data files\config.ini", "SONIDO", "VolM")) + Val(10))
                    Case 74
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "FPS") = "64 FPS" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "FPS", "FPS_MAX")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "FP", "23")
                    Else
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "FPS", "64 FPS")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "FP", "30")
                    End If
                    Case 75
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "FPS") = "64 FPS" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "FPS", "FPS_MAX")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "FP", "23")
                    Else
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "FPS", "64 FPS")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "FP", "30")
                    End If
                    
                    Case 76
                    'Nuevo
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "MODE") = "Windowed" Then Else Exit Sub
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = "1024×720" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "800")
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "600")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "23")
                    Exit Sub
                    End If
                    
                    ix = Screen.Width \ Screen.TwipsPerPixelX
                    iy = Screen.Height \ Screen.TwipsPerPixelY
                    
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = ix & "×" & iy Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "1024")
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "720")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Exit Sub
                    End If
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = "800×600" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "" & ix)
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "" & iy)
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Exit Sub
                    End If
                    
                    Case 77
                    '800×600 - 1024×720 - 1280×720 - 1366 ×720
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "MODE") = "Windowed" Then Else Exit Sub
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = "800×600" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "1024")
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "720")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Exit Sub
                    End If
                    
                    ix = Screen.Width \ Screen.TwipsPerPixelX
                    iy = Screen.Height \ Screen.TwipsPerPixelY
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = "1024×720" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "" & ix)
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "" & iy)
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Exit Sub
                    End If
                    
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH") & "×" & GetVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT") = ix & "×" & iy Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "800")
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "600")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "23")
                    Exit Sub
                    End If
                    
                    Case 78
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "MODE") = "Windowed" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "MODE", "Fullscreen")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "MP", "21")
                    ix = Screen.Width \ Screen.TwipsPerPixelX
                    iy = Screen.Height \ Screen.TwipsPerPixelY
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "" & ix)
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "" & iy)
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Else
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "MODE", "Windowed")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "MP", "21")
                    End If
                    
                    Case 79
                    If GetVar(App.Path & "\data files\config.ini", "Resolucion", "MODE") = "Windowed" Then
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "MODE", "Fullscreen")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "MP", "21")
                    ix = Screen.Width \ Screen.TwipsPerPixelX
                    iy = Screen.Height \ Screen.TwipsPerPixelY
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENWIDTH", "" & ix)
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "SCREENHEIGHT", "" & iy)
                    Call PutVar(App.Path & "\data files\config.ini", "System", "RP", "20")
                    Else
                    Call PutVar(App.Path & "\data files\config.ini", "Resolucion", "MODE", "Windowed")
                    Call PutVar(App.Path & "\data files\config.ini", "System", "MP", "21")
                    End If
                    
                    Case 72
                    If VisibleTextMenu = 6 Then Else Exit Sub
                     If OpcionClickeada = "1" Then
                     VisibleTextMenu = 4
                     OpcionClickeada = 0
                     Else
                     OpcionClickeada = "1"
                     End If
                     
                    Case 73
                    If VisibleTextMenu = 6 Then Else Exit Sub
                     If OpcionClickeada = "2" Then
                     OpcionClickeada = 0
                     VisibleTextMenu = 5
                     Else
                     OpcionClickeada = "2"
                     End If
                    Case 66
                        DestroyTCP
                        Alphamenu = 0
                        frmMenu.picCredits.Visible = False
                        Show_Login Not frmMenu.txtLUser.Visible
                        Show_Register False
                        Show_Character False
                        If VisibleTextMenu = 1 Then
                            VisibleTextMenu = 0
                            Buttons(70).Visible = False
                        Else
                            VisibleTextMenu = 1
                            Buttons(70).Visible = True
                        End If
                        
                        If frmMenu.txtLUser.Visible Then
                            frmMenu.txtLPass.SetFocus
                            frmMenu.txtLPass.SelStart = Len(frmMenu.txtLPass.text)
                        End If
                        ' Sonido
                        PlaySound Sound_ButtonClick, -1, -1
                    Exit Sub
                    
                    Case 67
                        DestroyTCP
                        Alphamenu = 0
                        frmMenu.picCredits.Visible = False
                        Show_Login False
                        Show_Register Not frmMenu.txtRUser.Visible
                        Show_Character False
                        If frmMenu.txtRUser.Visible Then
                            frmMenu.txtRUser.SetFocus
                        End If
                        If VisibleTextMenu = 2 Then
                            VisibleTextMenu = 0
                            Buttons(70).Visible = False
                        Else
                            VisibleTextMenu = 2
                            Buttons(70).Visible = True
                        End If
                        
                        ' play sound
                        PlaySound Sound_ButtonClick, -1, -1
                    Exit Sub
                    
                    Case 68
                        DestroyTCP
                        Alphamenu = 0
                        'picCredits.Visible = Not picCredits.Visible
                        Show_Login False
                        Show_Register False
                        Show_Character False
                        
                         
                        If VisibleTextMenu = 4 Then
                         VisibleTextMenu = 0
                         Buttons(70).Visible = False
                         PlaySound Sound_ButtonClick, -1, -1
                        Else
                         VisibleTextMenu = 4
                         Buttons(70).Visible = True
                        End If
                    Exit Sub
                    
                    Case 69
                        Call DestroyGame
                    Exit Sub
                    
                    Case 70
                    If VisibleTextMenu = 1 Then
                        If isLoginLegal(frmMenu.txtLUser.text, frmMenu.txtLPass.text) Then
                            Call MenuState(MENU_STATE_LOGIN)
                        End If
                    ElseIf VisibleTextMenu = 2 Then
                    
                        name = Trim$(frmMenu.txtRUser.text)
                        Password = Trim$(frmMenu.txtRPass.text)
                        PasswordAgain = Trim$(frmMenu.txtRPass2.text)
                    
                        If isLoginLegal(name, Password) Then
                            If Password <> PasswordAgain Then
                                Call MsgBox("No coinciden las Claves.")
                                Exit Sub
                            End If
                    
                            If Not isStringLegal(name) Then
                                Exit Sub
                            End If
                    
                            Call MenuState(MENU_STATE_NEWACCOUNT)
                            VisibleTextMenu = 0
                            Buttons(70).Visible = False
                        
                        End If
                     ElseIf VisibleTextMenu = 5 Then
                     
                      'If frmMenu.ConfigCo1 = "64 FPS" Or frmMenu.ConfigCo1 = "+64 FPS" And frmMenu.ConfigCo2 = "800x600" And frmMenu.ConfigCo3 = "Windowed" Or frmMenu.ConfigCo3 = "FullScreen" Then
                      VisibleTextMenu = 1
                     ElseIf VisibleTextMenu = 4 Then
                     
                      'If frmMenu.ConfigCo1 = "64 FPS" Or frmMenu.ConfigCo1 = "+64 FPS" And frmMenu.ConfigCo2 = "800x600" And frmMenu.ConfigCo3 = "Windowed" Or frmMenu.ConfigCo3 = "FullScreen" Then
                      VisibleTextMenu = 1
                      'rodo
                      'VTemp = Split(frmMenu.ConfigCo2, "×")
                      'Call PutVar(App.Path & "/data files/config.ini", "Resolucion", "FPS", frmMenu.ConfigCo1)
                      'Call PutVar(App.Path & "/data files/config.ini", "Resolucion", "SCREENWIDTH", VTemp(0))
                      'Call PutVar(App.Path & "/data files/config.ini", "Resolucion", "SCREENHEIGHT", VTemp(1))
                      'Call PutVar(App.Path & "/data files/config.ini", "Resolucion", "MODE", frmMenu.ConfigCo3)
                      'Call LoadDX8Vars
                    ElseIf VisibleTextMenu = 3 Then
                        Call MenuState(MENU_STATE_ADDCHAR)
                    End If
                End Select
            End If
        End If
    Next
End Sub
' scroll bar
Public Sub ChatScroll_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    ' find out which button we're clicking
    For I = 34 To 35
        x = GUIWindow(GUI_CHAT).x + Buttons(I).x
        y = GUIWindow(GUI_CHAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            ' scroll the actual chat
            Select Case I
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(I).x
        y = GUIWindow(GUI_SHOP).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 23
                        ' exit
                        Set buffer = New clsBuffer
                        buffer.WriteLong CCloseShop
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        GUIWindow(GUI_SHOP).Visible = False
                        InShop = 0
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub Achievement_MouseDown()
Dim I As Long, x As Long, y As Long

    ' find out which button we're clicking
    For I = 80 To 81
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'Buttons(I).state = 2 ' clicked
        If I = 80 Then
        If PaginaLogro = 1 Then Exit Sub
        PaginaLogro = PaginaLogro - 1
        End If
        
        If I = 81 Then
        If NombreLogro(((PaginaLogro + 1) * 4) - 3) = "" Then Exit Sub
        PaginaLogro = PaginaLogro + 1
        End If
        
        End If
    Next
End Sub
Public Sub Shop_MouseDown()
Dim I As Long, x As Long, y As Long

    ' find out which button we're clicking
    For I = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(I).x
        y = GUIWindow(GUI_SHOP).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub
Public Sub Bank_DoubleClick()
Dim bankNum As Long
    bankNum = IsBankItem(GlobalX, GlobalY)
    If bankNum <> 0 Then
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        If Item(GetBankItemNum(bankNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetBankItemNum(bankNum)).Stackable > 0 Then
            CurrencyMenu = 3 ' withdraw
            CurrencyText = "Que cantidad deseas tomar?"
            tmpCurrencyItem = bankNum
            sDialogue = vbNullString
            GUIWindow(GUI_CURRENCY).Visible = True
            inChat = True
            chatOn = True
            Exit Sub
        End If
        WithdrawItem bankNum, 0
        Exit Sub
    End If
End Sub
Public Sub Trade_DoubleClick()
Dim tradeNum As Long
    tradeNum = IsTradeItem(GlobalX, GlobalY, True)
    If tradeNum <> 0 Then
        UntradeItem tradeNum
        Exit Sub
    End If
End Sub
Public Sub Trade_MouseDown()
Dim I As Long, x As Long, y As Long

    ' find out which button we're clicking
    For I = 40 To 41
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Trade_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 40 To 41
        x = Buttons(I).x
        y = Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 40
                        AcceptTrade
                    Case 41
                        DeclineTrade
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Party
Public Sub Party_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(I).x
        y = GUIWindow(GUI_PARTY).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Objetivo inválido.", BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "No estás en un grupo.", BrightRed
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(I).x
        y = GUIWindow(GUI_PARTY).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Guild
Public Sub Guild_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 43 To 44
        x = GUIWindow(GUI_GUILD).x + Buttons(I).x
        y = GUIWindow(GUI_GUILD).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 43 ' Scroll Up
                        If GuildScroll > 1 Then GuildScroll = GuildScroll - 1
                    Case 44 ' Scroll Down
                        If GuildScroll < MAX_GUILD_MEMBERS - 4 And Not GuildData.Guild_Members(GuildScroll + 1).User_Name = vbNullString Then GuildScroll = GuildScroll + 1
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Guild_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 43 To 44
        x = GUIWindow(GUI_GUILD).x + Buttons(I).x
        y = GUIWindow(GUI_GUILD).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Combat
Public Sub Combat_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 52 To 53
        x = GUIWindow(GUI_COMBAT).x + Buttons(I).x
        y = GUIWindow(GUI_COMBAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 52 ' Scroll Up
                        If CombatScroll > 0 Then CombatScroll = CombatScroll - 1
                    Case 53 ' Scroll Down
                        'It's "- 4" beacuse you the first visible 4 don't count towards the scroll value
                        'You could also do: If CombatScroll + 4 < MAX_COMBAT + MAX_SKILLS Then
                        If CombatScroll < MAX_COMBAT + MAX_SKILLS - 4 Then CombatScroll = CombatScroll + 1
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Combat_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 52 To 53
        x = GUIWindow(GUI_COMBAT).x + Buttons(I).x
        y = GUIWindow(GUI_COMBAT).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For I = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case 26 ' music on
                        Options.Music = 1
                        PlayMusic Trim$(Map.Music)
                        SaveOptions
                        Buttons(26).state = 2
                        Buttons(27).state = 0
                    Case 27 ' music off
                        Options.Music = 0
                        StopMusic
                        SaveOptions
                        Buttons(26).state = 0
                        Buttons(27).state = 2
                    Case 28 ' sound on
                        Options.sound = 1
                        SaveOptions
                        Buttons(28).state = 2
                        Buttons(29).state = 0
                    Case 29 ' sound off
                        Options.sound = 0
                        StopAllSounds
                        SaveOptions
                        Buttons(28).state = 0
                        Buttons(29).state = 2
                    Case 30 ' debug on
                        Options.Debug = 1
                        SaveOptions
                        Buttons(30).state = 2
                        Buttons(31).state = 0
                    Case 31 ' debug off
                        Options.Debug = 0
                        SaveOptions
                        Buttons(30).state = 0
                        Buttons(31).state = 2
                    Case 32 ' levels on
                        Options.Lvls = 1
                        SaveOptions
                        Buttons(32).state = 2
                        Buttons(33).state = 0
                    Case 33 ' levels off
                        Options.Lvls = 0
                        SaveOptions
                        Buttons(32).state = 0
                        Buttons(33).state = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    For I = 59 To 62
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 3 Then
                ' do stuffs
                Select Case I
                    Case 59 ' FullScreen on
                        Buttons(59).state = 2
                        Buttons(60).state = 0
                        frmMain.BorderStyle = 0
                        'LoadDX8Vars
                        'StopRender = 0
                    Case 60 ' Pantalla completa off
                        PCD = 0
                        Buttons(59).state = 0
                        Buttons(60).state = 2
                    Case 61 ' buttons on
                        Options.Buttons = 1
                        SaveOptions
                        Buttons(61).state = 2
                        Buttons(62).state = 0
                    Case 62 ' buttons off
                        Options.Buttons = 0
                        SaveOptions
                        Buttons(61).state = 0
                        Buttons(62).state = 2
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next I
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
    
    For I = 59 To 62
        x = GUIWindow(GUI_OPTIONS).x + Buttons(I).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 0 Then
                Buttons(I).state = 3 ' clicked
            End If
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 1 To 7
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            If Buttons(I).state = 2 Then
                ' do stuffs
                Select Case I
                    Case 1
                        ' open window
                         OpenGuiWindow 1
                    Case 2
                        ' open window
                         OpenGuiWindow 2
                    Case 3
                        ' open window
                         OpenGuiWindow 3
                    Case 4
                        ' open window
                         OpenGuiWindow 4
                    Case 5
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendTradeRequest
                        Else
                            AddText "Objetivo inválido.", BrightRed
                        End If
                    Case 6
                        ' open window
                         OpenGuiWindow 6
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
    Next
    
    ' guild button
    I = 42
    
    x = GUIWindow(GUI_MENU).x + Buttons(I).x
    y = GUIWindow(GUI_MENU).y + Buttons(I).y
    ' check if we're on the button (Eliminado - Generaba Bug)
    If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then

        If Buttons(I).state = 2 Then
            ' do stuffs
             OpenGuiWindow 7
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
End If
    
    I = 65 'misiones boton
    x = GUIWindow(GUI_MENU).x + Buttons(I).x
    y = GUIWindow(GUI_MENU).y + Buttons(I).y
    ' check if we're on the button
    If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then

        If Buttons(I).state = 2 Then
            ' do stuffs
             OpenGuiWindow 8
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
End If
    ' reset buttons
    resetClickedButtons
End Sub
Public Sub OptionsMenu_MouseDown(ByVal Button As Long)
Dim I As Long, x As Long, y As Long
Dim texto As String
    ' find out which button we're clicking
            ' check if we're on the button
            'MouseTester.Label3.Caption = GlobalX >= X
            'MouseTester.Label4.Caption = GlobalX <= X + Buttons(I).Width
            'MouseTester.Label5.Caption = GlobalY >= Y
            'MouseTester.Label6.Caption = GlobalY <= Y + Buttons(I).Height
    ' guild button
    I = 71
    If Buttons(I).Visible Then
        x = GUIWindow(GUI_MENUOPTIONS).x + 56
        y = GUIWindow(GUI_MENUOPTIONS).y + 67
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            logoutGame
        End If
    End If

I = 4 'boton mision EaSee Engine
    If Buttons(I).Visible Then
        x = GUIWindow(GUI_MENU).x - 107
        y = GUIWindow(GUI_MENU).y - 206
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
            GUIWindow(GUI_OPTIONS).Visible = True
        End If
    End If

End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 1 To 6
        If Buttons(I).Visible Then
            x = GUIWindow(GUI_MENU).x + Buttons(I).x
            y = GUIWindow(GUI_MENU).y + Buttons(I).y
            ' check if we're on the button
            If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
                Buttons(I).state = 2 ' clicked
            End If
        End If
    Next
    
    ' guild button
    I = 42
    If Buttons(I).Visible Then
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    End If

I = 65 'boton mision EaSee Engine
    If Buttons(I).Visible Then
        x = GUIWindow(GUI_MENU).x + Buttons(I).x
        y = GUIWindow(GUI_MENU).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    End If

End Sub

' HP/MP/SP GUI
Public Sub Bars_MouseDown(ByVal Button As Long)
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 63 To 64
        x = GUIWindow(GUI_BARS).x + Buttons(I).x
        y = GUIWindow(GUI_BARS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                Select Case I
                    Case 63 ' minimap
                        Buttons(I).state = 2 'clicked
                    Case 64 ' buttons
                        Buttons(I).state = 2 'clicked
                End Select
                ' play sound - NO. double beeping is driving me crazy lol
                'PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    'resetClickedButtons
End Sub

' HP/MP/SP GUI
Public Sub Bars_MouseUp()
Dim I As Long, x As Long, y As Long, buffer As clsBuffer

    ' find out which button we're clicking
    For I = 63 To 64
        x = GUIWindow(GUI_BARS).x + Buttons(I).x
        y = GUIWindow(GUI_BARS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            'If Buttons(I).state = 2 Then
                Select Case I
                    Case 63 ' minimap
                        Options.MiniMap = Abs(Not CBool(Options.MiniMap))
                        SaveOptions
                        Buttons(I).state = 0 'normal
                        
                        If CBool(Options.MiniMap) = True Then
                            Buttons(59).state = 2 ' normal
                            Buttons(60).state = 3 ' clicked
                        Else
                            Buttons(59).state = 3 ' clicked
                            Buttons(60).state = 2 ' normal
                        End If
                    Case 64 ' buttons
                        Options.Buttons = Abs(Not CBool(Options.Buttons))
                        SaveOptions
                        Buttons(I).state = 0 'normal
                        
                        If CBool(Options.Buttons) = True Then
                            Buttons(61).state = 2 ' normal
                            Buttons(62).state = 3 ' clicked
                        Else
                            Buttons(61).state = 3 ' clicked
                            Buttons(62).state = 2 ' normal
                        End If
                End Select
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX - 12, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        mouseClicked = False
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long


    invNum = IsInvItem(GlobalX - 12, GlobalY)


    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If
    End If
End Sub

Public Sub Inventory_SingleClick()
    Dim invNum As Long, Value As Long, multiplier As Double, I As Long
    
    ' Not if we're in a shop
    If InShop > 0 Then Exit Sub
    ' Not if in bank
    If InBank Then Exit Sub
    ' Not if in trade
    If InTrade > 0 Then Exit Sub
    
    'Check if selected
        invNum = IsInvItem(GlobalX - 12, GlobalY)
        If invNum > 0 Then
            Call SendCheckHighlightItem(invNum)
            Exit Sub
        End If
    
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, Value As Long, multiplier As Double, I As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX - 12, GlobalY)
    
    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 2 ' deposit
                CurrencyText = "Que cantidad quieres depositar?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).Visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For I = 1 To MAX_INV
                If TradeYourOffer(I).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)).Stackable > 0 Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(I).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(I).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                CurrencyMenu = 4 ' offer in trade
                CurrencyText = "Que cantidad deseas intercambiar?"
                tmpCurrencyItem = invNum
                sDialogue = vbNullString
                GUIWindow(GUI_CURRENCY).Visible = True
                inChat = True
                chatOn = True
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum) > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum)).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub
' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub

Public Sub Dialogue_Button_MouseDown(Index As Integer)
    ' call the handler
    dialogueHandler Index
    GUIWindow(GUI_DIALOGUE).Visible = False
    inChat = False
    dialogueIndex = 0
End Sub

Public Sub Character_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 16 To 20
       x = GUIWindow(GUI_CHARACTER).x + Buttons(I).x + 50
 y = GUIWindow(GUI_CHARACTER).y + Buttons(I).y + 49
 If I > 18 Then x = GUIWindow(GUI_CHARACTER).x + 75 + Buttons(I).x
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 16 To 20
        x = GUIWindow(GUI_CHARACTER).x + Buttons(I).x + 50
        y = GUIWindow(GUI_CHARACTER).y + Buttons(I).y + 49
         If I > 18 Then x = GUIWindow(GUI_CHARACTER).x + 75 + Buttons(I).x
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (I - 15)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
    
    'equipando item canido


Dim eqNum As Long


eqNum = IsEqItem(GlobalX, GlobalY)


    If DragInvSlotNum > 0 Then
    Select Case Item(GetPlayerInvItemNum(MyIndex, DragInvSlotNum)).Type
    Case 1 'Arma
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 57 And GlobalX < GUIWindow(GUI_CHARACTER).x + 92 And GlobalY > GUIWindow(GUI_CHARACTER).y + 112 And GlobalY < GUIWindow(GUI_CHARACTER).y + 146 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 2 'Armadura
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 197 And GlobalX < GUIWindow(GUI_CHARACTER).x + 231 And GlobalY > GUIWindow(GUI_CHARACTER).y + 77 And GlobalY < GUIWindow(GUI_CHARACTER).y + 111 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 3 'Casco
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 197 And GlobalX < GUIWindow(GUI_CHARACTER).x + 231 And GlobalY > GUIWindow(GUI_CHARACTER).y + 42 And GlobalY < GUIWindow(GUI_CHARACTER).y + 76 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 4 'Pantalones
     If GlobalX > GUIWindow(GUI_CHARACTER).x + 197 And GlobalX < GUIWindow(GUI_CHARACTER).x + 231 And GlobalY > GUIWindow(GUI_CHARACTER).y + 112 And GlobalY < GUIWindow(GUI_CHARACTER).y + 146 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 5 'Botas
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 127 And GlobalX < GUIWindow(GUI_CHARACTER).x + 161 And GlobalY > GUIWindow(GUI_CHARACTER).y + 147 And GlobalY < GUIWindow(GUI_CHARACTER).y + 181 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 6 'Guantes
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 162 And GlobalX < GUIWindow(GUI_CHARACTER).x + 196 And GlobalY > GUIWindow(GUI_CHARACTER).y + 147 And GlobalY < GUIWindow(GUI_CHARACTER).y + 181 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 7 'Anillos
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 92 And GlobalX < GUIWindow(GUI_CHARACTER).x + 126 And GlobalY > GUIWindow(GUI_CHARACTER).y + 147 And GlobalY < GUIWindow(GUI_CHARACTER).y + 181 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 8 'Collar
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 57 And GlobalX < GUIWindow(GUI_CHARACTER).x + 92 And GlobalY > GUIWindow(GUI_CHARACTER).y + 42 And GlobalY < GUIWindow(GUI_CHARACTER).y + 76 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    Case 9 'Escudo
    If GlobalX > GUIWindow(GUI_CHARACTER).x + 57 And GlobalX < GUIWindow(GUI_CHARACTER).x + 92 And GlobalY > GUIWindow(GUI_CHARACTER).y + 77 And GlobalY < GUIWindow(GUI_CHARACTER).y + 111 Then
    Call SendUseItem(DragInvSlotNum)
    End If
    
End Select
End If
End Sub
' Npc Chat
Public Sub Chat_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long

If chatOnlyContinue = False Then
    For I = 1 To 4
        If Len(Trim$(chatOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
            x = GUIWindow(GUI_EVENTCHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_EVENTCHAT).y + 115 - ((I - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                chatOptState(I) = 2 ' clicked
            End If
        End If
    Next
Else
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    y = GUIWindow(GUI_EVENTCHAT).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        chatContinueState = 2 ' clicked
    End If
End If

End Sub
Public Sub Chat_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long

If chatOnlyContinue = False Then
    For I = 1 To 4
        If Len(Trim$(chatOpt(I))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(I)) & "]")
            x = GUIWindow(GUI_EVENTCHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_EVENTCHAT).y + 115 - ((I - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' are we clicked?
                If chatOptState(I) = 2 Then
                    SendChatOption I
                    ' play sound
                    PlaySound Sound_ButtonClick, -1, -1
                End If
            End If
        End If
    Next
    
    For I = 1 To 4
        chatOptState(I) = 0 ' normal
    Next
Else
    Width = EngineGetTextWidth(Font_Default, "[Continue]")
    x = GUIWindow(GUI_EVENTCHAT).x + ((GUIWindow(GUI_EVENTCHAT).Width / 2) - Width / 2)
    y = GUIWindow(GUI_EVENTCHAT).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        ' are we clicked?
        If chatContinueState = 2 Then
            SendChatContinue
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    
    chatContinueState = 0
End If
End Sub
Public Sub HandleKeyUp(ByVal KeyCode As Long)
Dim I As Long

    Select Case KeyCode
        Case vbKeyInsert
        Dim abierto As Boolean
        
             If Player(MyIndex).Access > 0 Then
             
                frmMain.picAdmin.Visible = Not frmMain.picAdmin.Visible
            End If
            If abierto = True Then
            chatOn = True
            abierto = False
            Else
            abierto = True
            chatOn = False 'EaSee fixChat 1
            End If
            
    End Select
    
    ' hotbar
    If Not chatOn And inChat = False Then
        For I = 1 To 9
            If KeyCode = 48 + I Then
                SendHotbarUse I
            End If
        Next
        If KeyCode = 48 Then ' 0
            SendHotbarUse 10
        ElseIf KeyCode = 189 Then ' -
            SendHotbarUse 11
        ElseIf KeyCode = 187 Then ' =
            SendHotbarUse 12
        End If
    End If
    
    ' handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If
End Sub
Public Sub QuestLog_MouseDown()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 45 To 51
        x = GUIWindow(GUI_QUESTLOG).x + Buttons(I).x
        y = GUIWindow(GUI_QUESTLOG).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next
End Sub
Public Sub Buddies_MouseDown()
Dim x As Long, y As Long, I As Long
    For I = 54 To 55
        x = GUIWindow(GUI_FRIENDS).x + Buttons(I).x
        y = GUIWindow(GUI_FRIENDS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 2 ' clicked
        End If
    Next I
End Sub
Public Sub FriendRequest_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If FriendRequestVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_FRIENDREQUEST).y + 80
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            FriendRequestAcceptState = 2 ' clicked
        End If
    End If
    
    Width = EngineGetTextWidth(Font_Georgia, "[Decline]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestDeclineState = 2 ' clicked
    End If
End Sub
Public Sub FriendRequest_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If Not FriendRequestVisible Then Exit Sub
    
    Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 80
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestAcceptState = 0 ' clicked
        ' Send accept packet to server
        Call SendAcceptFriend(FriendRequestSender)
        GUIWindow(GUI_FRIENDREQUEST).Visible = False
        FriendRequestVisible = False
        FriendRequestSender = vbNullString
    End If
    
    Width = EngineGetTextWidth(Font_Georgia, "[Decline]")
    x = (GUIWindow(GUI_FRIENDREQUEST).x + (GUIWindow(GUI_FRIENDREQUEST).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_FRIENDREQUEST).y + 100
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        FriendRequestDeclineState = 0 ' clicked
        ' Send decline packet to server
        Call SendDeclineFriend(FriendRequestSender)
        GUIWindow(GUI_FRIENDREQUEST).Visible = False
        FriendRequestVisible = False
        FriendRequestSender = vbNullString
    End If
End Sub
Public Sub Book_MouseDown()
Dim x As Long, y As Long, I As Long
Dim Parse() As String
    For I = 56 To 58
        x = GUIWindow(GUI_BOOK).x + Buttons(I).x
        y = GUIWindow(GUI_BOOK).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 0 ' normal
            
            Select Case I
                ' Left button
                Case 56
                    Book_PageLeft = True
                    Exit Sub
                ' Right button
                Case 57
                    Book_PageRight = True
                    Exit Sub
                ' X button
                Case 58
                    GUIWindow(GUI_BOOK).Visible = False
                    OpeningBook = True
                    Exit Sub
            End Select
        End If
    Next I
End Sub
Public Sub PlayerInfo_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Georgia, "[X]")
    x = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width - 25))
    y = GUIWindow(GUI_FRIENDS).y + 10
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        PlayerInfoX = 2 ' clicked
    End If
End Sub
Public Sub PlayerInfo_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    Width = EngineGetTextWidth(Font_Georgia, "[X]")
    x = (GUIWindow(GUI_FRIENDS).x + (GUIWindow(GUI_FRIENDS).Width - 25))
    y = GUIWindow(GUI_FRIENDS).y + 10
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        PlayerInfoX = 0 ' clicked
        GUIWindow(GUI_PLAYERINFO).Visible = False
        GUIWindow(GUI_FRIENDS).Visible = True
        frmMain.lstFriends.Visible = True
    End If
End Sub
Public Sub Buddies_MouseUp()
Dim x As Long, y As Long, I As Long
Dim Parse() As String
    For I = 54 To 55
        x = GUIWindow(GUI_FRIENDS).x + Buttons(I).x
        y = GUIWindow(GUI_FRIENDS).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            Buttons(I).state = 0 ' normal
            
            Select Case I
                ' Defriend button
                Case 54
                    'is something selected and if so, does it have text?
                    If frmMain.lstFriends.ListIndex < 0 Then Exit Sub
                    If Not Len(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)) > 0 Then Exit Sub
                    
                    ' Delete friend
                    SendDeleteFriend frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)
                    Exit Sub
                ' Message button
                Case 55
                    If frmMain.lstFriends.ListIndex < 0 Then Exit Sub ' Nothing selected
                    If Not Len(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex)) > 0 Then Exit Sub 'No name in selection
                    If InStr(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex), "Offline") > 0 Then Exit Sub ' Player is offline

                    chatOn = True
                    Parse() = Split(frmMain.lstFriends.List(frmMain.lstFriends.ListIndex), " ")
                    MyText = "!" & Parse(0) & " "
                    RenderChatText = MyText
                    Exit Sub
            End Select
        End If
    Next I
End Sub
Public Sub QuestLog_MouseUp()
Dim I As Long, x As Long, y As Long
    ' find out which button we're clicking
    For I = 45 To 51
        x = GUIWindow(GUI_QUESTLOG).x + Buttons(I).x
        y = GUIWindow(GUI_QUESTLOG).y + Buttons(I).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(I).Width) And (GlobalY >= y And GlobalY <= y + Buttons(I).Height) Then
            ' send the level up
            If Trim$(frmMain.lstQuestLog.text) = vbNullString Then Exit Sub
                LoadQuestlogBox (I - 44)
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    Next
End Sub
Public Sub QuestAccept_MouseDown()
    PlayerHandleQuest CLng(QuestAcceptTag), 1
    inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).Visible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
    RefreshQuestLog
End Sub
Public Sub QuestExtra_MouseDown()
    RunQuestDialogueExtraLabel
End Sub

Public Sub QuestClose_MouseDown()
    inChat = False
    GUIWindow(GUI_QUESTDIALOGUE).Visible = False
    
    QuestExtraVisible = False
    QuestAcceptVisible = False
    QuestAcceptTag = vbNullString
    QuestSay = "-"
End Sub
Public Sub QuestDialogue_MouseDown()
Dim I As Long, x As Long, y As Long, Width As Long
    
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            QuestAcceptState = 2 ' clicked
        End If
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 120
        'If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            QuestExtraState = 2 ' clicked
        End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_CHAT).y + 120
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        QuestCloseState = 2 ' clicked
    End If
End Sub

Public Sub QuestDialogue_MouseUp()
Dim I As Long, x As Long, y As Long, Width As Long
    If QuestAcceptVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[Accept]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 105
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If QuestAcceptState = 2 Then
                QuestAccept_MouseDown
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        QuestAcceptState = 0
    End If
    If QuestExtraVisible = True Then
        Width = EngineGetTextWidth(Font_Georgia, "[" & QuestExtra & "]")
        x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
        y = GUIWindow(GUI_CHAT).y + 120
        If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
            If QuestExtraState = 2 Then
                QuestExtra_MouseDown
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            End If
        End If
        QuestExtraState = 0
    End If
    Width = EngineGetTextWidth(Font_Georgia, "[Close]")
    x = (GUIWindow(GUI_QUESTDIALOGUE).x + (GUIWindow(GUI_QUESTDIALOGUE).Width / 2)) - (Width / 2)
    y = GUIWindow(GUI_CHAT).y + 120
    If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
        If QuestCloseState = 2 Then
            QuestClose_MouseDown
            ' play sound
            PlaySound Sound_ButtonClick, -1, -1
        End If
    End If
    QuestCloseState = 0
End Sub
