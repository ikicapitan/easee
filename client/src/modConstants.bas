Attribute VB_Name = "modConstants"
Option Explicit

' API Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByRef MSG() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' *******************************************************************************
' ***************************---3 FRAMED MOVEMENT---*****************************
' Set to true if you want VX, VX ace style 3 framed movement
' If its set to false, you will have XP style 4 framed movement
Public Const VXFRAME As Boolean = False
' *******************************************************************************

' animated buttons
Public Const MAX_MENUBUTTONS As Long = 4
Public Const MAX_MAINBUTTONS As Long = 8
Public Const MENUBUTTON_PATH As String = "\Data Files\graphics\gui\menu\buttons\"
Public Const MAINBUTTON_PATH As String = "\Data Files\graphics\gui\main\buttons\"

' Skills
Public Const MAX_SKILLS_INIT As Long = 4
Public MAX_SKILLS As Long

' Friends
Public Const MAX_FRIENDS As Long = 50

' Animation
Public Const AnimColumns As Long = 5

' Attack Time Without weapon
Public Const AttTimeNoWep As Long = 750

' Hotbar
Public Const HotbarTop As Long = 2
Public Const HotbarLeft As Long = 2
Public Const HotbarOffsetX As Long = 9

' Inventory constants
Public Const InvTop As Long = 4
Public Const InvLeft As Long = 10
Public Const InvOffsetY As Long = 3
Public Const InvOffsetX As Long = 3
Public Const InvColumns As Long = 7

' Bank constants
Public Const BankTop As Long = 38
Public Const BankLeft As Long = 42
Public Const BankOffsetY As Long = 3
Public Const BankOffsetX As Long = 4
Public Const BankColumns As Long = 11

' spells constants
Public Const SpellTop As Long = 4
Public Const SpellLeft As Long = 10
Public Const SpellOffsetY As Long = 3
Public Const SpellOffsetX As Long = 3
Public Const SpellColumns As Long = 5

' shop constants
Public Const ShopTop As Long = 24
Public Const ShopLeft As Long = 38
Public Const ShopOffsetY As Long = 3
Public Const ShopOffsetX As Long = 3
Public Const ShopColumns As Long = 5

' Character consts
Public Const EqTop As Long = 164
Public Const EqLeft As Long = 43
Public Const EqOffsetX As Long = 3
Public Const EqOffsetY As Long = 4
Public Const EqColumns As Long = 3

' values
Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

' path constants
Public Const SOUND_PATH As String = "\Data Files\sound\"
Public Const MUSIC_PATH As String = "\Data Files\music\"

' Font variables
Public Const FONT_NAME As String = "Georgia"
Public Const FONT_SIZE As Byte = 14

' Log Path and variables
Public Const LOG_DEBUG As String = "debug.txt"
Public Const LOG_PATH As String = "\Data Files\logs\"

' Map Path and variables
Public Const MAP_PATH As String = "\data files\maps\"
Public Const MAP_EXT As String = ".map"

' Gfx Path and variables
Public Const GFX_PATH As String = "\data files\graphics\"
Public Const GFX_EXT As String = ".png"

Public Const FONT_PATH As String = "\data files\graphics\fonts\"

' Key constants
Public Const VK_UP As Long = &H26
Public Const VK_DOWN As Long = &H28
Public Const VK_LEFT As Long = &H25
Public Const VK_RIGHT As Long = &H27
Public Const VK_SHIFT As Long = &H10
Public Const VK_RETURN As Long = &HD
Public Const VK_CONTROL As Long = &H11

' Menu states
Public Const MENU_STATE_NEWACCOUNT As Byte = 0
Public Const MENU_STATE_DELACCOUNT As Byte = 1
Public Const MENU_STATE_LOGIN As Byte = 2
Public Const MENU_STATE_GETCHARS As Byte = 3
Public Const MENU_STATE_NEWCHAR As Byte = 4
Public Const MENU_STATE_ADDCHAR As Byte = 5
Public Const MENU_STATE_DELCHAR As Byte = 6
Public Const MENU_STATE_USECHAR As Byte = 7
Public Const MENU_STATE_INIT As Byte = 8

' Speed moving vars
Public WALK_SPEED As Long
Public RUN_SPEED As Long
Public Const PARALIZADO As Long = 0 'Para futuros scripts y funciones de EaSee
' Tile size constants
Public Const PIC_X As Long = 32
Public Const PIC_Y As Long = 32

' Sprite, item, spell size constants
Public Const SIZE_X As Long = 39
Public Const SIZE_Y As Long = 39



' ********************************************************
' * The values below must match with the server's values *
' ********************************************************

' General constants
Public Const MAX_PLAYERS As Long = 70
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 100
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 12
Public Const MAX_PARTYS As Long = 35
Public Const MAX_PARTY_MEMBERS As Long = 4
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000
Public Const MAX_WEATHER_PARTICLES As Long = 250
Public Const MAX_COMBAT As Byte = 8
Public Const MAX_COMBAT_LEVEL As Byte = 100
Public Const MAX_COMBO As Long = 100
Public Const MAX_COMBO_GIVEN As Long = 10
Public Const MAX_PLAYER_PROJECTILES As Long = 1000 'Cantidad de flechas a disparar por jugador en cada sesion. Constante requerida.

' Website
Public Const GAME_WEBSITE As String = "http://www.easee.es"

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16
Public Const Orange As Byte = 17

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const MUSIC_LENGTH As Byte = 40
Public Const ACCOUNT_LENGTH As Byte = 12
Public Const SKILL_LENGTH As Byte = 30
Public Const TEXT_LENGTH As Long = 1000

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 100
Public MAX_MAPX As Byte '= 24
Public MAX_MAPY As Byte '= 18
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1

' GUI
Public Const MAX_BUTTONS As Long = 89

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_KEY As Byte = 5
Public Const TILE_TYPE_KEYOPEN As Byte = 6
Public Const TILE_TYPE_RESOURCE As Byte = 7
Public Const TILE_TYPE_DOOR As Byte = 8
Public Const TILE_TYPE_NPCSPAWN As Byte = 9
Public Const TILE_TYPE_SHOP As Byte = 10
Public Const TILE_TYPE_BANK As Byte = 11
Public Const TILE_TYPE_HEAL As Byte = 12
Public Const TILE_TYPE_TRAP As Byte = 13
Public Const TILE_TYPE_SLIDE As Byte = 14
Public Const TILE_TYPE_SOUND As Byte = 15
Public Const TILE_TYPE_PLAYERSPAWN As Byte = 16
Public Const TILE_TYPE_LETRERO As Byte = 17

'Weather Type Constants
Public Const WEATHER_TYPE_NONE As Byte = 0
Public Const WEATHER_TYPE_RAIN As Byte = 1
Public Const WEATHER_TYPE_SNOW As Byte = 2
Public Const WEATHER_TYPE_HAIL As Byte = 3
Public Const WEATHER_TYPE_SANDSTORM As Byte = 4
Public Const WEATHER_TYPE_STORM As Byte = 5

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_HELMET As Byte = 3
Public Const ITEM_TYPE_LEGS As Byte = 4 ' New
Public Const ITEM_TYPE_BOOTS As Byte = 5 ' New
Public Const ITEM_TYPE_GLOVE As Byte = 6 ' New
Public Const ITEM_TYPE_RING As Byte = 7 ' New
Public Const ITEM_TYPE_ENCHANT As Byte = 8 ' New
Public Const ITEM_TYPE_SHIELD As Byte = 9
Public Const ITEM_TYPE_CONSUME As Byte = 10
Public Const ITEM_TYPE_KEY As Byte = 11
Public Const ITEM_TYPE_CURRENCY As Byte = 12
Public Const ITEM_TYPE_SPELL As Byte = 13
Public Const ITEM_TYPE_BOOK As Byte = 14
Public Const ITEM_TYPE_CUBO As Byte = 15
Public Const ITEM_TYPE_PICACUBOS As Byte = 16
Public Const ITEM_TYPE_MUNICION As Byte = 17


' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3

' Constants for player movement: Tiles per Second
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const MAX_NPC_DROP_ITEMS As Byte = 10
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_DAMAGEHP As Byte = 0
Public Const SPELL_TYPE_DAMAGEMP As Byte = 1
Public Const SPELL_TYPE_HEALHP As Byte = 2
Public Const SPELL_TYPE_HEALMP As Byte = 3
Public Const SPELL_TYPE_WARP As Byte = 4
Public Const SPELL_TYPE_BUFF As Byte = 5 'EaSee 0.5

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4
Public Const EDITOR_RESOURCE As Byte = 5
Public Const EDITOR_ANIMATION As Byte = 6
Public Const EDITOR_COMBO As Boolean = 7

' Target type constants
Public Const TARGET_TYPE_NONE As Byte = 0
Public Const TARGET_TYPE_PLAYER As Byte = 1
Public Const TARGET_TYPE_NPC As Byte = 2
Public Const TARGET_TYPE_EVENT As Byte = 4

' Dialogue box constants
Public Const DIALOGUE_TYPE_NONE As Byte = 0
Public Const DIALOGUE_TYPE_TRADE As Byte = 1
Public Const DIALOGUE_TYPE_FORGET As Byte = 2
Public Const DIALOGUE_TYPE_PARTY As Byte = 3

' Do Events
Public Const nLng As Long = (&H80 Or &H1 Or &H4 Or &H20) + (&H8 Or &H40)

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2

' stuffs
Public HalfX As Integer  '= ((MAX_MAPX + 1) / 2) * PIC_X
Public HalfY As Integer  '= ((MAX_MAPY + 1) / 2) * PIC_Y
Public ScreenX As Integer  '= (MAX_MAPX + 1) * PIC_X
Public ScreenY As Integer  '= (MAX_MAPY + 1) * PIC_Y
Public StartXValue As Integer  '= ((MAX_MAPX + 1) / 2)
Public StartYValue As Integer  '= ((MAX_MAPY + 1) / 2)
Public EndXValue As Integer  '= (MAX_MAPX + 1) + 1
Public EndYValue As Integer  '= (MAX_MAPY + 1) + 1
Public Half_PIC_X As Integer  '= PIC_X / 2
Public Half_PIC_Y As Integer  '= PIC_Y / 2

' Autotiles
Public Const AUTO_INNER As Byte = 1
Public Const AUTO_OUTER As Byte = 2
Public Const AUTO_HORIZONTAL As Byte = 3
Public Const AUTO_VERTICAL As Byte = 4
Public Const AUTO_FILL As Byte = 5

' Autotile types
Public Const AUTOTILE_NONE As Byte = 0
Public Const AUTOTILE_NORMAL As Byte = 1
Public Const AUTOTILE_FAKE As Byte = 2
Public Const AUTOTILE_ANIM As Byte = 3
Public Const AUTOTILE_CLIFF As Byte = 4
Public Const AUTOTILE_WATERFALL As Byte = 5

' Rendering
Public Const RENDER_STATE_NONE As Long = 0
Public Const RENDER_STATE_NORMAL As Long = 1
Public Const RENDER_STATE_AUTOTILE As Long = 2

'Chatbubble
Public Const ChatBubbleWidth As Long = 200

Public Const EFFECT_TYPE_FADEIN As Long = 1
Public Const EFFECT_TYPE_FADEOUT As Long = 2
Public Const EFFECT_TYPE_FLASH As Long = 3
Public Const EFFECT_TYPE_FOG As Long = 4
Public Const EFFECT_TYPE_WEATHER As Long = 5
Public Const EFFECT_TYPE_TINT As Long = 6

'EaSee Engine

Public cuboactualizar As Boolean 'Actualiza cubos en cada actualizacion de items
Public procesandocubo As Boolean
Public procesandopicacubos As Boolean
Public enfriarveneno As Boolean
Public FBP As Boolean
Public FBN As Boolean
Public DBP As Boolean
Public DBN As Boolean
Public ABP As Boolean
Public ABN As Boolean
Public IBP As Boolean
Public IBN As Boolean
Public VBP As Boolean
Public VBN As Boolean
Public SpriteH As Boolean
Public TiempoIntro As Long
Public TiempoConstanteIntro As Long
Public FramesWaitIntro As Long 'Y por que en ingles chabon? xD
Public Count_Surface As Long
Public FinIntro As Boolean
Public ColorIntro As Long
Public PosicionIntro2 As Long
Public IntroFase1 As Boolean 'Flag
Public IntroMusica As Boolean 'Flag

