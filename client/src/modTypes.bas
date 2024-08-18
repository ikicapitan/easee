Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map As MapRec
Public Bank As BankRec
Public TempTile() As TempTileRec
Public Player(1 To MAX_PLAYERS) As PlayerRec
Public Class() As ClassRec
Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAP_NPCS) As MapNpcRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapSounds() As MapSoundRec
Public MapSoundCount As Long
Public WeatherParticle(1 To MAX_WEATHER_PARTICLES) As WeatherParticleRec
Public Autotile() As AutotileRec
Public Skill() As SkillDataRec
' client-side stuff
Public ActionMsg(1 To MAX_BYTE) As ActionMsgRec
Public Blood(1 To MAX_BYTE) As BloodRec
Public AnimInstance(1 To MAX_BYTE) As AnimInstanceRec
Public MenuButton(1 To MAX_MENUBUTTONS) As OLD_ButtonRec
Public MainButton(1 To MAX_MAINBUTTONS) As OLD_ButtonRec
Public Party As PartyRec
Public GUIWindow() As GUIWindowRec
Public Buttons(1 To MAX_BUTTONS) As ButtonRec
Public Chat(1 To 20) As ChatRec
Public Combo(1 To MAX_COMBO) As ComboRec
Public SlotLibro As Long

Public Type BookRec
    name As String * NAME_LENGTH
    text As String * TEXT_LENGTH
    Text2 As String * TEXT_LENGTH
End Type

'combo system
Private Type ComboRec
    Item_1 As Long
    Item_2 As Long
    Item_Given(1 To MAX_COMBO_GIVEN) As Long
    Item_Given_Val(1 To MAX_COMBO_GIVEN) As Long
    Skill As Long
    SkillLevel As Long
    Level As Long
    Take_Item1 As Byte
    Take_Item2 As Byte
    GiveSkill As Long
    GiveSkill_Exp As Long
    ReqItem1 As Long
    ReqItem2 As Long
    ReqItemVal1 As Long
    ReqItemVal2 As Long
    Take_ReqItem1 As Byte
    Take_ReqItem2 As Byte
End Type

' skills
Private Type SkillDataRec
    name As String * SKILL_LENGTH
    MaxLvl As Long
    Div As Long
End Type

' options
Public Options As OptionsRec
'Evilbunnie's DrawnChat system
Private Type ChatRec
    text As String
    colour As Long
End Type

' Type recs
Private Type OptionsRec
    Game_Name As String
    Resol_Ancho As Long
    Resol_Alto As Long
    savePass As Byte
    Password As String * NAME_LENGTH
    Username As String * ACCOUNT_LENGTH
    IP As String
    Port As Long
    IntroMusic As String
    MenuMusic As String
    Music As Byte
    sound As Byte
    Debug As Byte
    Lvls As Byte
    MiniMap As Byte
    Buttons As Byte
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerInvRec
    num As Long
    Value As Long
    Selected As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerInvRec
End Type

Private Type SpellAnim
    spellnum As Long
    Timer As Long
    FramePointer As Long
End Type

Private Type CombatRec
    Level As Byte
    EXP As Long
End Type

Private Type PlayerSpawnRec
    Map As Byte
    X As Byte
    Y As Byte
End Type

Private Type SkillRec
    Level As Long
    EXP As Long
    EXP_Needed As Long
End Type

Private Type FriendRec
    NameOfFriend(1 To MAX_FRIENDS) As String * NAME_LENGTH
    Count As Long
End Type

Public Type ProjectileRec
    TravelTime As Long
    Direction As Long
    X As Long
    Y As Long
    Pic As Long
    Range As Long
    Damage As Long
    speed As Long
    Municion As Long
End Type

Private Type PlayerRec
    ' General
    name As String
    Class As Long
    Sprite As Long
    Level As Byte
    EXP As Long
    Access As Byte
    PK As Byte
    ' Admins
    Visible As Long
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    MaxVital(1 To Vitals.Vital_Count - 1) As Long
    ' Stats
    Stat(1 To Stats.Stat_Count - 1) As Byte
    POINTS As Long
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As Long
    ' Position
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    PlayerQuest(1 To MAX_QUESTS) As PlayerQuestRec
    ' Spawn location
    Spawn As PlayerSpawnRec
    ' Combat
    Combat(1 To MAX_COMBAT) As CombatRec
    ' Client use only
    xOffset As Integer
    yOffset As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    Step As Byte
    EventTimer As Long
    StartFlash As Long
    StartVeneno As Long 'EaSee 0.6
    StartParalisis As Long 'EaSee 0.6
    ' Guild
    GuildName As String
    GuildTag As String
    GuildColor As Integer
    
    'NEW Everything below is new
    'Walkthrough
    Walkthrough As Boolean
    
    'Follow feature
    Follower As Long
    
    'Skills
    Skills() As SkillRec
    
    'Friends
    Friends As FriendRec
    
    'Projectiles
    ProjecTile(1 To MAX_PLAYER_PROJECTILES) As ProjectileRec

    VCorrer As Long 'EaSee 0.3
    VCaminar As Long
    Oculto As Boolean 'EaSee 0.4
    FuerzaBuff As Long 'EaSee 0.5
    DestrezaBuff As Long
    AgilidadBuff As Long
    InteligenciaBuff As Long
    VoluntadBuff As Long
    FuerzaDeBuff As Long 'EaSee 0.5
    DestrezaDeBuff As Long
    AgilidadDeBuff As Long
    InteligenciaDeBuff As Long
    VoluntadDeBuff As Long

End Type

Private Type TileDataRec
    X As Long
    Y As Long
    Tileset As Long
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Public Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
End Type

Public Type EventCommandRec
    Index As Long
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    Data5 As Long
    Data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Public Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Public Type EventPageRec
    'These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    HasItemAmount As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    'End Conditions
    
    'Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    'Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    'Guidelines for the event
    WalkAnim As Byte
    DirFix As Byte
    Walkthrough As Byte
    ShowName As Byte
    
    'Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    'Client Needed Only
    X As Long
    Y As Long
End Type

Public Type EventRec
    name As String
    Global As Long
    pageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
End Type

Public Type TileRec
    layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
    Cubo As Boolean
    HP As Long 'Vida del Cubo
    Mensaje As String 'Letreros
    Animacion As Long 'Animacion
    
    'Para futuras versiones
    Banco As Long
    Evento As Long
    BancoLlave As Long
    Script As Long
    Timer As Long
    ParteCubo As Boolean
    SFX1 As Long
    SFX2 As Long
    Objeto As Long
    
End Type

Private Type MapEventRec
    name As String
    Dir As Long
    X As Long
    Y As Long
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    Moving As Long
    MovementSpeed As Long
    Position As Long
    xOffset As Long
    yOffset As Long
    Step As Long
    Visible As Long
    WalkAnim As Long
    DirFix As Long
    ShowDir As Long
    Walkthrough As Long
    ShowName As Long
End Type

Private Type MapRec
    name As String * NAME_LENGTH
    Music As String * MUSIC_LENGTH
    BGS As String * MUSIC_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Red As Long
    Green As Long
    Blue As Long
    alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
    DropItemsOnDeath As Byte
    Cubo As Boolean
    
    'Client Side Only -- Temporary
    CurrentEvents As Long
    MapEvents() As MapEventRec
End Type

Private Type ClassRec
    name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_Count - 1) As Byte
    MaleSprite() As Long
    FemaleSprite() As Long
    Oculto() As Boolean
    VCaminar() As Long
    VCorrer() As Long
    Faccion() As Long 'EaSee '0.6
    ItemFaccion As Long
    ' For client use
    Vital(1 To Vitals.Vital_Count - 1) As Long
End Type

Private Type ItemRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    CombatTypeReq As Byte
    CombatLvlReq As Byte
    
    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    ClassReq As Long
    AccessReq As Long
    LevelReq As Long
    Mastery As Byte
    Price As Long
    Add_Stat(1 To Stats.Stat_Count - 1) As Byte
    Rarity As Byte
    speed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_Count - 1) As Byte
    Animation As Long
    Paperdoll As Long
    
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    instaCast As Byte
    Stackable As Byte
    
    'NEW Everything below is new
    SkillReq As Long
    
    'Datos Item Cubo (EaSee Engine)
    CuboTileX As Long 'variables de datos Cubo EaSee
    CuboTileY As Long
    CuboTileN As Long
    Cubo32 As Boolean
    Cubo64 As Boolean
    CuboCapa1 As Byte
    CuboCapa2 As Byte
    CuboSupTipo As Byte '1 para bloqueo, 2 para cofre, se pueden agregar mas
    CuboInfTipo As Byte '1 para bloqueo, 2 para cofre
    CuboMapa As Integer 'Para cubo de transportacion
    CuboMapaX As Integer
    CuboMapaY As Integer
    CuboGolpe As Integer 'Para cubo trampa
    CuboDureza As Integer 'Vida del Cubo
    CuboAnimacion As Integer 'Animacion del cubo al destruirse
    CuboSFX1 As Integer 'Sonido al golpear
    CuboSFX2 As Integer 'Sonido al Destruir
    CuboObjeto As Integer
    
    
    'Elements
    Element_Light_Dmg As Long
    Element_Light_Res As Long
    Element_Dark_Dmg As Long
    Element_Dark_Res As Long
    Element_Neut_Dmg As Long
    Element_Neut_Res As Long
    
    'books
    Book As BookRec
    
    'Projectiles
    ProjecTile As ProjectileRec
End Type

Private Type MapItemRec
    PlayerName As String
    num As Long
    Value As Long
    Frame As Byte
    X As Byte
    Y As Byte
End Type

Private Type NpcDropRec
    DropChance As Long
    DropItem As Long
    DropItemValue As Long
    RandCurrency As Byte
    P_5 As Long
    P_10 As Long
    P_20 As Long
End Type

Private Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    sound As String * NAME_LENGTH
    
    Sprite As Long
    SpawnSecs As Long
    Behaviour As Byte
    Range As Byte
    Drops(1 To MAX_NPC_DROP_ITEMS) As NpcDropRec
    Stat(1 To Stats.Stat_Count - 1) As Byte
    HP As Long
    EXP As Long
    Animation As Long
    Damage As Long
    Level As Long
    speed As Long
    Quest As Byte
    QuestNum As Long
    
    'NEW Everything below is new
    RandExp As Byte
    Percent_5 As Byte
    Percent_10 As Byte
    Percent_20 As Byte
    RandHP As Byte
    HPMin As Long
    Element_Light_Dmg As Long
    Element_Light_Res As Long
    Element_Dark_Dmg As Long
    Element_Dark_Res As Long
    Element_Neut_Dmg As Long
    Element_Neut_Res As Long
    RndSpawn As Byte
    SpawnSecsMin As Long
End Type

Private Type MapNpcRec
    num As Long
    target As Long
    targetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    HPSetTo As Long
    ' Client use only
    xOffset As Long
    yOffset As Long
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    Step As Byte
    StartFlash As Long
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
End Type

Private Type ShopRec
    name As String * NAME_LENGTH
    BuyRate As Long
    TradeItem(1 To MAX_TRADES) As TradeItemRec
End Type

Private Type SpellRec
    name As String * NAME_LENGTH
    Desc As String * 255
    sound As String * NAME_LENGTH
    CombatTypeReq As Byte
    CombatLvlReq As Byte
    
    Type As Long
    MPCost As Long
    LevelReq As Long
    AccessReq As Long
    ClassReq As Long
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    IsAoE As Boolean
    AoE As Long
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    
    'NEW
    Dmg_Light As Long
    Dmg_Dark As Long
    Dmg_Neut As Long
    Paralisis As Long 'EaSee Engine 0.5
    Inversion As Long
    Invisibilidad As Long
    Veneno As Long 'Usa el Valor HP para Quitar
    Velocidad As Long 'Agregar o Quitar
    VelocidadCaminar2 As Long
    VelocidadCorrer2 As Long
    Buff As Long '0 para Debuff 1 para Buff
    Fuerza As Long
    Destreza As Long
    Agilidad As Long
    Inteligencia As Long
    Voluntad As Long
    Necesidad1 As Long 'Necesidades creables 0.5
    Necesidad2 As Long
    Necesidad3 As Long
    Necesidad4 As Long
    Arco As Long
    NumeroArcoItem As Long
    Sprite As Long
    NumeroSprite As Long
    Clase As Long
    NumeroClase As Long
    Transportar As Long
    TransportarX As Long
    TransportarY As Long
    TransportarMapa As Long
    Desequipar As Long
    VenenoDmg As Long
End Type

Private Type TempTileRec
    DoorOpen As Byte
    DoorFrame As Byte
    DoorTimer As Long
    DoorAnimate As Byte ' 0 = nothing| 1 = opening | 2 = closing
End Type

Public Type MapResourceRec
    X As Long
    Y As Long
    ResourceState As Byte
End Type

Private Type ResourceRec
    name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    ResourceType As Byte
    ResourceImage As Long
    ExhaustedImage As Long
    ItemReward As Long
    ToolRequired As Long
    health As Long
    RespawnTime As Long
    Walkthrough As Boolean
    Animation As Long
    
    'NEW Everything below is new
    ItemRewardAmount As Long
    healthmin As Long
    HPRand As Byte
    DistItems As Byte
    ItemRewardAmountMin As Long
    ItemRewardRand As Byte
    
    'Skill stuff
    Exp_Give As Byte
    Exp_Amnt As Long
    Exp_Skill As Long
    Skill_Req As Long
    Skill_LvlReq As Long
    
    'Msg Colors
    Color_Success As Long
    Color_Empty As Long
End Type

Private Type ActionMsgRec
    Message As String
    Created As Long
    Type As Long
    Color As Long
    Scroll As Long
    X As Long
    Y As Long
    Timer As Long
End Type

Private Type BloodRec
    Sprite As Long
    Timer As Long
    X As Long
    Y As Long
End Type

Private Type AnimationRec
    name As String * NAME_LENGTH
    sound As String * NAME_LENGTH
    
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    looptime(0 To 1) As Long
End Type

Private Type AnimInstanceRec
    Animation As Long
    X As Long
    Y As Long
    ' used for locking to players/npcs
    lockindex As Long
    LockType As Byte
    ' timing
    Timer(0 To 1) As Long
    ' rendering check
    Used(0 To 1) As Boolean
    ' counting the loop
    LoopIndex(0 To 1) As Long
    frameIndex(0 To 1) As Long
End Type

Public Type HotbarRec
    Slot As Long
    sType As Byte
End Type

Public Type OLD_ButtonRec
    filename As String
    state As Byte
End Type
Public Type ButtonRec
    state As Byte
    X As Long
    Y As Long
    Width As Long
    Height As Long
    Visible As Boolean
    PicNum As Long
End Type

Public Type GUIWindowRec
    X As Long
    Y As Long
    Width As Long
    Height As Long
    Visible As Boolean
End Type

Public Type TexGuiRec
    X As Long
    Y As Long
    Width As Long
    Height As Long
End Type

Public Type EventListRec
    CommandList As Long
    CommandNum As Long
End Type

Public Type MapSoundRec
    X As Long
    Y As Long
    SoundHandle As Long
    InUse As Boolean
    Channel As Long
End Type

Public Type WeatherParticleRec
    Type As Long
    X As Long
    Y As Long
    Velocity As Long
    InUse As Long
End Type

'Auto tiles "/
Public Type PointRec
    X As Long
    Y As Long
End Type

Public Type QuarterTileRec
    QuarterTile(1 To 4) As PointRec
    RenderState As Byte
    srcX(1 To 4) As Long
    srcY(1 To 4) As Long
End Type

Public Type AutotileRec
    layer(1 To MapLayer.Layer_Count - 1) As QuarterTileRec
End Type

Public Type ChatBubbleRec
    Msg As String
    colour As Long
    target As Long
    targetType As Byte
    Timer As Long
    active As Boolean
End Type

' Mini Map Data
Public MiniMapPlayer(1 To MAX_PLAYERS) As MiniMapPlayerRec
Public MiniMapNPC(1 To MAX_MAP_NPCS) As MiniMapNPCRec

Public Type MiniMapPlayerRec
    X As Long
    Y As Long
End Type

Public Type MiniMapNPCRec
    X As Long
    Y As Long
End Type
