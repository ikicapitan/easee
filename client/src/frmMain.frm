VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   7740
   ClientLeft      =   4125
   ClientTop       =   1635
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   753
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton ClanBoton 
      Caption         =   "Fundar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   52
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox ClanTag 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9720
      MaxLength       =   3
      TabIndex        =   51
      Top             =   5760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox ClanName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9360
      MaxLength       =   10
      TabIndex        =   50
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrdebuff 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   3240
   End
   Begin VB.Timer tmrHechizos 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   5400
   End
   Begin VB.Timer tmrScrollEditor 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   4680
   End
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00B5B5B5&
      ForeColor       =   &H80000008&
      Height          =   5730
      Left            =   360
      Picture         =   "frmMain.frx":4FEA
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   347
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   5235
      Begin VB.CommandButton Command1 
         Caption         =   "X"
         Height          =   255
         Left            =   4800
         TabIndex        =   49
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton btnWalkthrough 
         Caption         =   "Pasar Atraves"
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         ToolTipText     =   "Permite que los jugadores se puedan atravesar (de esta forma no se bloquearan el paso en caminos angostos por ejemplo)"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdAKill 
         Caption         =   "Asesinar"
         Height          =   255
         Left            =   1680
         TabIndex        =   34
         ToolTipText     =   "Efectua la muerte del jugador indicado."
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAHeal 
         Caption         =   "Curar"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         ToolTipText     =   "Restaura la VIDA HP del jugador."
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAName 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         ToolTipText     =   "Modifica el nombre del jugador."
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdLevel 
         Caption         =   "Subir Nivel"
         Height          =   255
         Left            =   3000
         TabIndex        =   31
         ToolTipText     =   "Sube 1 nivel al personaje indicado."
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animación"
         Height          =   255
         Left            =   3360
         TabIndex        =   30
         ToolTipText     =   $"frmMain.frx":1DBA4
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Acceso"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         ToolTipText     =   "Modifica los privilegios del jugador."
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtAAccess 
         Height          =   285
         Left            =   1680
         TabIndex        =   28
         ToolTipText     =   "Escriba el numero de privilegio a cambiar."
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtASprite 
         Height          =   285
         Left            =   4440
         TabIndex        =   27
         ToolTipText     =   "Escriba el numero de Sprite."
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         ToolTipText     =   "ReActualiza o Reenvia el Mapa actual (lo carga nuevamente)."
         Top             =   1845
         Width           =   855
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Sprite"
         Height          =   255
         Left            =   4920
         TabIndex        =   25
         ToolTipText     =   "Modifica la imagen de tu personaje (Sprite) por la que has indicado en numero."
         Top             =   5400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Insertar Objeto"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         ToolTipText     =   "Inserta el Objeto especificado en el mapa actual."
         Top             =   5400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   4560
         Min             =   1
         TabIndex        =   23
         Top             =   5400
         Value           =   1
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   4560
         Min             =   1
         TabIndex        =   22
         Top             =   5400
         Value           =   1
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Hechizos"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         ToolTipText     =   $"frmMain.frx":1DC76
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Tienda"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         ToolTipText     =   $"frmMain.frx":1DD51
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Recurso"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         ToolTipText     =   $"frmMain.frx":1DDD8
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         ToolTipText     =   "Abre el Editor de NPC para la creacion o edicion de los mismos (monstruos del juego, personajes no jugadores, etc)"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Mapa"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         ToolTipText     =   "Abre el Editor de Mapas para diseñar las pantallas o niveles de tu juego asignando atributos, imagenes, eventos, NPC y demas."
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Objetos"
         Height          =   255
         Left            =   3360
         TabIndex        =   16
         ToolTipText     =   $"frmMain.frx":1DEAA
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdADestroy 
         Caption         =   "Desbanear"
         Height          =   255
         Left            =   1680
         TabIndex        =   15
         ToolTipText     =   "Quita el Ban a un Jugador especifico."
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Reportar Mapa"
         Height          =   225
         Left            =   3000
         TabIndex        =   14
         ToolTipText     =   "Genera un Reporte del Mapa Actual."
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Coord"
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         ToolTipText     =   "Exhibe las coordenadas X, Y y el Mapa en el que se encuentra el Personaje y el Cursor."
         Top             =   1845
         Width           =   855
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Ir a"
         Height          =   195
         Left            =   3000
         TabIndex        =   12
         ToolTipText     =   "Te transporta al Mapa que has indicado."
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txtAMap 
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Escriba el numero de mapa."
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton cmdAWarpMe2 
         Caption         =   "Ir Hacia"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         ToolTipText     =   "Te transporta hacia el jugador."
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdAWarp2Me 
         Caption         =   "Traer A"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         ToolTipText     =   "Transporta al jugador hacia tu sitio."
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Banear"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         ToolTipText     =   "Expulsa de forma permanente al jugador del juego."
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kickear"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         ToolTipText     =   "Expulsa al jugador del juego."
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtAName 
         Height          =   285
         Left            =   480
         TabIndex        =   6
         ToolTipText     =   "Escriba el nombre del personaje."
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdAVisible 
         Caption         =   "Visible"
         Height          =   255
         Left            =   480
         TabIndex        =   5
         ToolTipText     =   "Te hace visible/invisible para el resto de los jugadores."
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdACharEdit 
         Caption         =   "Personaje"
         Height          =   255
         Left            =   2040
         TabIndex        =   4
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdAQuest 
         Caption         =   "Misiones"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         ToolTipText     =   "Abre el Editor de Misiones o Quest, donde podras crear o editar misiones otorgables por NPC para ser realizadas por los jugadores."
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite#:"
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad: 1"
         Height          =   255
         Left            =   4920
         TabIndex        =   36
         Top             =   5400
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Insertar Objeto: Ninguno"
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   5400
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.ListBox lstQuestLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      ItemData        =   "frmMain.frx":1DF36
      Left            =   3360
      List            =   "frmMain.frx":1DF38
      MousePointer    =   1  'Arrow
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   17280
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   9960
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstFriends 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2280
      ItemData        =   "frmMain.frx":1DF3A
      Left            =   600
      List            =   "frmMain.frx":1DF3C
      MousePointer    =   1  'Arrow
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame frmMensaje 
      BackColor       =   &H00004090&
      Caption         =   "Letrero"
      ForeColor       =   &H8000000B&
      Height          =   3615
      Left            =   1440
      TabIndex        =   46
      Top             =   2640
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   435
         Left            =   1440
         TabIndex        =   47
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label lblTexto 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   1815
         Left            =   480
         TabIndex        =   48
         Top             =   600
         Width           =   3615
      End
   End
   Begin VB.Frame frmeditorletrero 
      BackColor       =   &H00004090&
      Caption         =   "Letrero"
      ForeColor       =   &H8000000B&
      Height          =   3615
      Left            =   1440
      TabIndex        =   40
      Top             =   2640
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   435
         Left            =   2280
         TabIndex        =   45
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Colocar"
         Height          =   435
         Left            =   840
         TabIndex        =   44
         Top             =   3000
         Width           =   1335
      End
      Begin VB.ComboBox cmbtxtcolor 
         Height          =   315
         ItemData        =   "frmMain.frx":1DF3E
         Left            =   2520
         List            =   "frmMain.frx":1DF40
         TabIndex        =   43
         Text            =   "Color"
         Top             =   2340
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.HScrollBar scrllletra 
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   2400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin RichTextLib.RichTextBox txteditorletrero 
         Height          =   1815
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3201
         _Version        =   393217
         TextRTF         =   $"frmMain.frx":1DF42
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ************
' ** Events **
' ************

Private MoveForm As Boolean
Private MouseX As Long
Private MouseY As Long
Private PresentX As Long
Private PresentY As Long
Private mouseClicked As Boolean


Private Sub btnWalkthrough_Click()
    SendWalkthrough
End Sub

Private Sub ClanBoton_Click()
Call GuildMake(frmMain.ClanName, frmMain.ClanTag)
End Sub

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub

    SendRequestEditAnimation
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAAnim_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdACharEdit_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    
    frmEditor_Character.Visible = True
End Sub

Private Sub cmdAHeal_Click()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    If Len(Trim$(txtAName.text)) > 2 Then
        SendHealPlayer Trim$(txtAName.text)
    Else
        If Len(txtAName.text) = 0 Then SendHealPlayer GetPlayerName(MyIndex)
    End If
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAHeal_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAKick_Click()
If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub
SendKick Trim$(txtAName.text)
End Sub

Private Sub cmdAKill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then Exit Sub

    If Len(Trim$(txtAName.text)) < 2 Then Exit Sub

    If IsNumeric(Trim$(txtAName.text)) Then Exit Sub

    SendKillPlayer Trim$(txtAName.text)

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAKill_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAName_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) < 2 Then
    Exit Sub
    End If
    
    If IsNumeric(Trim$(txtAName.text)) Or IsNumeric(Trim$(txtAAccess.text)) Then
    Exit Sub
    End If
    
    SendSetName Trim$(txtAName.text), (Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAName_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAQuest_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditQuest
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAVisible_Click()
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then Exit Sub
    SendVisibility
End Sub

Private Sub cmdcancelar_Click()
frmeditorletrero.Visible = False
CanMoveNow = True 'Desbloquea PJ
End Sub

Private Sub cmdCerrar_Click()
frmMensaje.Visible = False
lblTexto.Caption = ""
End Sub

Private Sub cmdguardar_Click()
    'Cubo1
    Dim Objeto As Integer
    Dim x As Long
    Dim y As Long
    Dim TileX As Long
    Dim TileY As Long
    Dim tilenum As Long
    Dim CuboSupTipo, CuboInfTipo As Byte
    Dim Mapa As Integer
    Dim Dato, Dato2, Dato3 As Byte
    Dim letrero As Boolean
    Dim Data4, Mensaje, Mensaje2 As String 'Lo usaremos tambien para el letrero
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
    Animacion = 0 'Para uso futuro
    Evento = 0
    Banco = 0
    BancoLlave = 0
    Script = 0
    Timer = 0
    SFX1 = Item(Objeto).CuboSFX1
    SFX2 = Item(Objeto).CuboSFX2
    SFX01 = Item(Objeto).CuboSFX1
    SFX02 = Item(Objeto).CuboSFX2 'Hasta aqui uso futuro
    Dropeo = Item(Objeto).CuboObjeto

Objeto = GetPlayerEquipment(MyIndex, Weapon)

    If Objeto > 0 Then
    If Item(Objeto).Type = ITEM_TYPE_CUBO Then 'Si hay un cubo equipado
        
    'Tomamos datos base
    x = GetPlayerX(MyIndex) 'Toma coordenadas del jugador
    y = GetPlayerY(MyIndex)
    CuboSupTipo = Item(Objeto).CuboSupTipo 'Atributo de Cubo
    Mapa = GetPlayerMap(MyIndex)
    
    Select Case GetPlayerDir(MyIndex) 'Toma la direccion en que mira el personaje para insertar
        Case DIR_UP
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y - 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 1 'Si es Bloqueo
        MsgBox ("Debes equiparte el Cubo adecuado")
    
        Case 2 'Banco
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 3 'Transporte
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 4 'Trampa
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 5 'Mensaje
        If y > 1 Then
        Map.Tile(x, y - 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y - 1).Type, Mapa, x, y - 1, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)
        frmeditorletrero.Visible = False
        End If
        End Select
        
                     
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
        Dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y - 2).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        Map.Tile(x, y - 2).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
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
        Map.Tile(x, y - 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje2 = txteditorletrero.text
        
        End Select
                
        
        If y > 1 Then  'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
         Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y - 2).Type, Mapa, x, y - 2, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x, y - 1).Type, x, y - 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje2)
        frmeditorletrero.Visible = False
        End If
        
        
        End If
        
        
        
        
        
        Case DIR_DOWN
        
        
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x, y + 1).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 1 'Si es Bloqueo
        MsgBox ("Debes equiparte el Cubo adecuado")
    
        Case 2 'Banco
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 3 'Transporte
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 4 'Trampa
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 5 'Mensaje
        If y > 1 Then
        Map.Tile(x, y + 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y - 1).Type, Mapa, x, y - 1, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)
        frmeditorletrero.Visible = False
        End If
        End Select
        
                     
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
        Dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x, y + 2).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        Map.Tile(x, y + 2).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
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
        Map.Tile(x, y + 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje2 = txteditorletrero.text
        
        End Select
                
        
        If y > 1 Then  'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
       Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x, y + 2).Type, Mapa, x, y + 1, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x, y + 1).Type, x, y + 2, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje2)
        frmeditorletrero.Visible = False
        End If
        
        
        End If
        
        
        
        
                    
                    
        Case DIR_LEFT
            
            
            
        
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 1 'Si es Bloqueo
        MsgBox ("Debes equiparte el Cubo adecuado")
    
        Case 2 'Banco
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 3 'Transporte
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 4 'Trampa
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 5 'Mensaje
        If y > 1 Then
        Map.Tile(x - 1, y).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y).Type, Mapa, x - 1, y, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)
        frmeditorletrero.Visible = False
        End If
        End Select
        
                     
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
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_BANK  'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        Map.Tile(x - 1, y + 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        End Select
                
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x - 1, y).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
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
        Map.Tile(x - 1, y).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje2 = txteditorletrero.text
        
        End Select
                
        
        If y > 1 Then  'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y + 1).Type, Mapa, x - 1, y, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x - 1, y).Type, x - 1, y + 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje2)
        frmeditorletrero.Visible = False
        End If
        
        
        End If
            
            
            
        Case DIR_RIGHT
            
            
        If Item(Objeto).Cubo32 = True Then 'Funcion direccion arriba para Cubo de 32x32
        
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).Tileset = Item(Objeto).CuboTileN 'Numero Tile
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa1).y = Item(Objeto).CuboTileY 'Tile Y
        
        Select Case CuboSupTipo
        
        Case 0 'Sin Atributo
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 1 'Si es Bloqueo
        MsgBox ("Debes equiparte el Cubo adecuado")
    
        Case 2 'Banco
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 3 'Transporte
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 4 'Trampa
        MsgBox ("Debes equiparte el Cubo adecuado")
        
        Case 5 'Mensaje
        If y > 1 Then
        Map.Tile(x + 1, y).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        Call EnviarMapaCubos(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y).Type, Mapa, x + 1, y, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo)
        frmeditorletrero.Visible = False
        End If
        End Select
        
                     
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
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_BANK  'Dibuja Banco o Cofre
        
        Case 3 'Transporte
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_WARP 'Dibuja Transporte
        Dato = Item(Objeto).CuboMapa
        Dato2 = Item(Objeto).CuboMapaX
        Dato3 = Item(Objeto).CuboMapaY
        Data4 = ""
        
        Case 4 'Trampa
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_TRAP 'Dibuja Trampa
        Dato = Item(Objeto).CuboGolpe
        Dato2 = 0
        Dato3 = 0
        Data4 = ""
        
        Case 5 'Mensaje
        Map.Tile(x + 1, y + 1).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje = txteditorletrero.text
        End Select
                
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa2).x = Item(Objeto).CuboTileX 'Tile X
        Map.Tile(x + 1, y).layer(Item(Objeto).CuboCapa2).y = Item(Objeto).CuboTileY + 1 'Tile Y
                
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
        Map.Tile(x + 1, y).Type = TILE_TYPE_LETRERO 'Parametros base
        Mensaje2 = txteditorletrero.text
        
        End Select
                
        
        If y > 1 Then  'solo es posible si la posicion del jugador es mayor a 1Y por el limite de pantalla, ni tampoco es un letrero
        Call EnviarMapaCubos64(Item(Objeto).CuboTileN, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y + 1).Type, Mapa, x + 1, y, Dato, Dato2, Dato3, Data4, HP, Animacion, Banco, Evento, BancoLlave, Script, Timer, SFX1, SFX2, Mensaje, Dropeo, Item(Objeto).CuboTileX, Item(Objeto).CuboTileY + 1, Item(Objeto).CuboCapa1, Map.Tile(x + 1, y).Type, x + 1, y + 1, Dato02, Dato22, Dato32, Data42, HP2, Animacion2, Banco2, Evento2, BancoLlave2, Script2, Timer2, SFX01, SFX02, Mensaje2)
        frmeditorletrero.Visible = False
        End If
        
        
        End If
            
            
            
            
            
    End Select

    
    End If
End If
End If
End If
End Sub

Private Sub cmdLevel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        SendRequestLevelUp GetPlayerName(MyIndex)
    Else
        SendRequestLevelUp Trim$(txtAName.text)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdLevel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub





Private Sub Command1_Click()
frmMain.picAdmin.Visible = False
End Sub

Private Sub Command2_Click()
OpenGuiWindow 11
End Sub

'Private Sub Form_Click()
'    HandleSingleClick
'End Sub

Private Sub Form_DblClick()
    HandleDoubleClick
End Sub
Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "AdminPanel", "L5")
lblAItem.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "L6")
lblAAmount = trad

trad = GetVar(App.Path & Lang, "AdminPanel", "B1")
cmdAKick.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B2")
cmdABan.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B3")
cmdAWarp2Me.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B4")
cmdAWarpMe2.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B5")
cmdAAccess.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B6")
cmdAName.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B7")
cmdAHeal.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B8")
cmdAKill.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B9")
cmdAVisible.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B10")
cmdAWarp.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B11")
cmdAMap.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B12")
cmdAItem.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B13")
cmdAResource.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B14")
cmdANpc.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B15")
cmdASpell.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B16")
cmdAShop.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B17")
cmdAAnim.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B18")
cmdACharEdit.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B19")
cmdAQuest.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B20")
btnWalkthrough.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B21")
cmdALoc.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B22")
cmdAMapReport.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B23")
cmdADestroy.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B24")
cmdARespawn.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B25")
cmdASpawn.Caption = trad
trad = GetVar(App.Path & Lang, "AdminPanel", "B26")
cmdLevel.Caption = trad


    ' move GUI
    picAdmin.Left = 444
    picAdmin.Top = 8
    mouseClicked = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseDown Button
End Sub

    Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim invNum As Long
invNum = DragInvSlotNum

If GlobalX < GUIWindow(GUI_INVENTORY).x And GlobalX < GUIWindow(GUI_INVENTORY).x + GUIWindow(GUI_INVENTORY).Width Then
If GlobalY > GUIWindow(GUI_INVENTORY).y And GlobalY < GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_INVENTORY).Height Then
 If invNum > 0 Then
    If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, invNum)).Stackable > 0 Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        CurrencyText = "Que cantidad deseas dejar?"
                        tmpCurrencyItem = invNum
                        sDialogue = vbNullString
                        GUIWindow(GUI_CURRENCY).Visible = True
                        inChat = True
                        chatOn = False 'EaSee Fix Chat 2
                    End If
                 End If
            End If
                         Call SendDropItem(invNum, 0)
   End If
 End If
                     
End If
End If
    HandleMouseUp Button
End Sub

Private Sub Form_Resize()
    If Not frmMain.Visible Then Exit Sub
     'Fullscreen work
    'LoadDX8Vars
    'InitDX8
    'GUIWindow(GUI_CHAT).Y = frmMain.ScaleHeight - 155
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Cancel = True
    logoutGame
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleMouseMove CLng(x), CLng(y), Button
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Function IsShopItem(ByVal x As Single, ByVal y As Single) As Long
Dim tempRec As RECT
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    IsShopItem = 0

    For I = 1 To MAX_TRADES

        If Shop(InShop).TradeItem(I).Item > 0 And Shop(InShop).TradeItem(I).Item <= MAX_ITEMS Then
            With tempRec
                .Top = ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If x >= tempRec.Left And x <= tempRec.Right Then
                If y >= tempRec.Top And y <= tempRec.Bottom Then
                    IsShopItem = I
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
ErrorHandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function







Public Sub lstFriends_DblClick()
Dim Parse() As String
    If lstFriends.ListIndex < 0 Then Exit Sub ' Nothing selected
    If Not Len(lstFriends.List(lstFriends.ListIndex)) > 0 Then Exit Sub 'No name in selection
    'If InStr(lstFriends.List(lstFriends.ListIndex), "Offline") > 0 Then Exit Sub ' Player is offline
    
    'This will load a gui for the player's data.
    Parse() = Split(lstFriends.List(lstFriends.ListIndex), " ")
    Call RequestFriendData(Parse(0))
End Sub



Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblAAmount.Caption = "Cantidad: " & scrlAAmount.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAAmount_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblAItem.Caption = "Objeto: " & Trim$(Item(scrlAItem.Value).name)
    If Item(scrlAItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlAItem.Value).Stackable > 0 Then
        scrlAAmount.Enabled = True
        Exit Sub
    End If
    scrlAAmount.Enabled = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAItem_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call HandleKeyPresses(KeyAscii)

    ' prevents textbox on error ding sound
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Then
        KeyAscii = 0
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim I As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    HandleKeyUp KeyCode

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Admin Menu **
' ****************

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    BLoc = Not BLoc
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdALoc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendRequestEditMap
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAMap_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp2Me_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpToMe Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarp2Me_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarpMe2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Then
        Exit Sub
    End If

    WarpMeTo Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarpMe2_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAWarp_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then
        Exit Sub
    End If

    If Not IsNumeric(Trim$(txtAMap.text)) Then
        Exit Sub
    End If

    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAWarp_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
    
    Exit Sub
    End If
    
    If Len(Trim$(txtASprite.text)) < 1 Then
    Exit Sub
    End If
    
    If Not IsNumeric(Trim$(txtASprite.text)) Then
    Exit Sub
    End If
    
    If Len(Trim$(txtAName.text)) > 1 Then
    SendSetSprite CLng(Trim$(txtASprite.text)), txtAName.text
    Else
    SendSetSprite CLng(Trim$(txtASprite.text)), GetPlayerName(MyIndex)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASprite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    SendMapReport
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAMapReport_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If
    
    SendMapRespawn
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdARespawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdABan_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 1 Then
        Exit Sub
    End If

    SendBan Trim$(txtAName.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdABan_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditItem
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditNpc
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdANpc_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditResource
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAResource_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditShop
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
        
        Exit Sub
    End If

    SendRequestEditSpell
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASpell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    If Len(Trim$(txtAName.text)) < 2 Then
        Exit Sub
    End If

    If IsNumeric(Trim$(txtAName.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then
        Exit Sub
    End If

    SendSetAccess Trim$(txtAName.text), CLng(Trim$(txtAAccess.text))
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdAAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdADestroy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
        
        Exit Sub
    End If

    SendBanDestroy
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdADestroy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then Exit Sub
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdASpawn_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picAdmin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseClicked = True
End Sub

Private Sub picAdmin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseClicked = False
End Sub

Private Sub picAdmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mouseClicked Then
        picAdmin.Left = picAdmin.Left + x
        picAdmin.Top = picAdmin.Top + y
    End If
End Sub

Private Sub Timer1_Timer()
    'RenderTexture Tex_MenuBg, GlobalX, GlobalY, GlobalX, GlobalY, 500, 500, 500, 500
    RenderTexture Tex_MenuBg2, 0, 0, PosMenu, 0, 800, 600, 800, 600
End Sub

Private Sub tmrdebuff_Timer()
If MyIndex = 0 Then Exit Sub
If VenenoDuracion = 0 Then Player(MyIndex).StartVeneno = 0
If ParalisisDuracion = 0 Then Player(MyIndex).StartParalisis = 0
If ConfusionDuracion < 1 And ParalisisDuracion < 1 And VenenoDuracion < 1 And VelocidadDuracion < 1 And InvisibilidadDuracion < 1 And FuerzaDuracionP < 1 And FuerzaDuracionN < 1 And DestrezaDuracionP < 1 And DestrezaDuracionN < 1 And AgilidadDuracionP < 1 And AgilidadDuracionN < 1 And InteligenciaDuracionP < 1 And InteligenciaDuracionN < 1 And VoluntadDuracionP < 1 And VoluntadDuracionN < 1 And SpriteDuracion < 1 Then
tmrdebuff.Enabled = False
Exit Sub
End If

If ConfusionDuracion > 0 Then 'Confusion
ConfusionDuracion = ConfusionDuracion - 1 'fin Confusion
End If

If ParalisisDuracion > 0 Then ParalisisDuracion = ParalisisDuracion - 1 'Paralisis
If VenenoDuracion > 0 Then 'Veneno
Call Golpe(VenenoGolpe)
VenenoDuracion = VenenoDuracion - 1

End If

If VelocidadDuracion > 0 Then VelocidadDuracion = VelocidadDuracion - 1 'Velocidad
If InvisibilidadDuracion > 0 Then InvisibilidadDuracion = InvisibilidadDuracion - 1 'Invisibilidad
If FuerzaDuracionP > 0 Then FuerzaDuracionP = FuerzaDuracionP - 1 'Buff Fza
If FuerzaDuracionN > 0 Then FuerzaDuracionN = FuerzaDuracionN - 1 'Debuff Fza
If DestrezaDuracionP > 0 Then DestrezaDuracionP = DestrezaDuracionP - 1 'Des
If DestrezaDuracionN > 0 Then DestrezaDuracionN = DestrezaDuracionN - 1
If AgilidadDuracionP > 0 Then AgilidadDuracionP = AgilidadDuracionP - 1 'Agi
If AgilidadDuracionN > 0 Then AgilidadDuracionN = AgilidadDuracionN - 1
If InteligenciaDuracionP > 0 Then InteligenciaDuracionP = InteligenciaDuracionP - 1 'Int
If InteligenciaDuracionN > 0 Then InteligenciaDuracionN = InteligenciaDuracionN - 1
If VoluntadDuracionP > 0 Then VoluntadDuracionP = VoluntadDuracionP - 1 'Vol
If VoluntadDuracionN > 0 Then VoluntadDuracionN = VoluntadDuracionN - 1
If SpriteDuracion > 0 Then SpriteDuracion = SpriteDuracion - 1 'Sprite
End Sub

Private Sub tmrScrollEditor_Timer()
    Scroll_Timer = Scroll_Timer + 1
    
    If Scroll_Timer >= 10 Then '1 second
        Select Case Scroll_Editor
            Case 1 'map editor
                If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo Continue
                SendRequestEditMap
            Case 2 'npc editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditNpc
            Case 3 'item editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditItem
            Case 4 'resource editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditResource
            Case 5 'quest editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditQuest
            Case 6 'spell editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditSpell
            Case 7 'character editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                frmEditor_Character.Visible = True
                frmEditor_Character.txtEName.text = GetPlayerName(MyIndex)
            Case 8 'animation editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditAnimation
            Case 9 'shop editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditShop
            Case 10 'combo editor
                If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo Continue
                SendRequestEditCombo
        End Select
Continue:
        
        Scroll_Timer = 0
        Scroll_Editor = 0
        Scroll_Draw = False
        tmrScrollEditor.Enabled = False
    End If
End Sub
