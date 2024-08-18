VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargando..."
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   503
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consola"
      TabPicture(0)   =   "frmServer.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCpsLock"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCPS"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtChat"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtText"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmestado"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tmrGetTime"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Cmdbrowser"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Jugadores"
      TabPicture(1)   =   "frmServer.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwInfo"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Control "
      TabPicture(2)   =   "frmServer.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraServer"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "fraDatabase"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "fraclases"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Gremios"
      TabPicture(3)   =   "frmServer.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdGSave"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Jugabilidad"
      TabPicture(4)   =   "frmServer.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "frmmaxmin"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "framaxymin"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame4"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Quartz"
      TabPicture(5)   =   "frmServer.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.CommandButton Cmdbrowser 
         Caption         =   "Comunidad"
         Height          =   315
         Left            =   8280
         TabIndex        =   98
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   855
         Left            =   -74280
         TabIndex        =   97
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Frame frmmaxmin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Maximos y Minimos"
         Height          =   4575
         Left            =   -68400
         TabIndex        =   93
         Top             =   360
         Width           =   3735
         Begin VB.CommandButton txtguardarmaxmin 
            Caption         =   "Guardar"
            Height          =   315
            Left            =   1200
            TabIndex        =   96
            Top             =   3960
            Width           =   1215
         End
         Begin VB.TextBox txtmaxclases 
            Height          =   285
            Left            =   840
            MaxLength       =   2
            TabIndex        =   94
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Clases:"
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Timer tmrGetTime 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7800
         Top             =   360
      End
      Begin VB.Frame frmestado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado del Juego"
         Height          =   3735
         Left            =   7680
         TabIndex        =   85
         Top             =   840
         Width           =   2775
         Begin VB.Label lbltiempo 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   495
            Left            =   480
            TabIndex        =   86
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame framaxymin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Central Ecologica del Tiempo"
         Height          =   4575
         Left            =   -71640
         TabIndex        =   77
         Top             =   360
         Width           =   3135
         Begin VB.TextBox txtopacidad 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   89
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox txtvelocidad 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   88
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox txtniebla 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   87
            Top             =   2160
            Width           =   855
         End
         Begin VB.TextBox txtintensidadclima 
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   81
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtmapaclima 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   80
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cmbclima 
            Height          =   315
            ItemData        =   "frmServer.frx":0972
            Left            =   120
            List            =   "frmServer.frx":0988
            TabIndex        =   79
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdclima 
            Caption         =   "Generar"
            Height          =   255
            Left            =   960
            TabIndex        =   78
            Top             =   3960
            Width           =   1215
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Opacidad"
            Height          =   255
            Left            =   1080
            TabIndex        =   92
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Velocidad"
            Height          =   255
            Left            =   1080
            TabIndex        =   91
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Niebla"
            Height          =   255
            Left            =   1080
            TabIndex        =   90
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Intensidad"
            Height          =   255
            Left            =   1080
            TabIndex        =   84
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lbltipoclima 
            BackStyle       =   0  'Transparent
            Caption         =   "Clima"
            Height          =   255
            Left            =   1440
            TabIndex        =   83
            Top             =   960
            Width           =   495
         End
         Begin VB.Label lblmapa 
            BackStyle       =   0  'Transparent
            Caption         =   "Numero Mapa (0 para Todos)"
            Height          =   495
            Left            =   1440
            TabIndex        =   82
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraclases 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clases"
         Height          =   4575
         Left            =   -69840
         TabIndex        =   47
         Top             =   480
         Width           =   5175
         Begin VB.CheckBox chkvisible 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oculta"
            Height          =   255
            Left            =   3600
            TabIndex        =   99
            Top             =   3000
            Width           =   855
         End
         Begin VB.TextBox txtcorrervelocidad 
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   76
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox txtcaminarvelocidad 
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   75
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox txtyspawn 
            Height          =   285
            Left            =   1200
            TabIndex        =   71
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox txtxspawn 
            Height          =   285
            Left            =   360
            TabIndex        =   70
            Top             =   3600
            Width           =   495
         End
         Begin VB.TextBox txtmapaspawn 
            Height          =   285
            Left            =   720
            TabIndex        =   69
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtvoluntadclase 
            Height          =   285
            Left            =   960
            TabIndex        =   63
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtagilidadclase 
            Height          =   285
            Left            =   3120
            TabIndex        =   62
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox txtresistenciaclase 
            Height          =   285
            Left            =   1200
            TabIndex        =   61
            Top             =   1800
            Width           =   975
         End
         Begin VB.TextBox txtinteligenciaclase 
            Height          =   285
            Left            =   3480
            TabIndex        =   60
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtfuerzaclase 
            Height          =   285
            Left            =   840
            TabIndex        =   59
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txtspritefemclase 
            Height          =   285
            Left            =   4080
            TabIndex        =   58
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtspritemascclase 
            Height          =   285
            Left            =   1680
            TabIndex        =   57
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtnombreclase 
            Height          =   285
            Left            =   960
            TabIndex        =   56
            Top             =   720
            Width           =   2775
         End
         Begin VB.CommandButton cmdguardarclases 
            Caption         =   "Guardar Clase"
            Height          =   255
            Left            =   1800
            TabIndex        =   55
            Top             =   4080
            Width           =   1695
         End
         Begin MSComctlLib.Slider sldclasenum 
            Height          =   375
            Left            =   1320
            TabIndex        =   54
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1
            Min             =   1
            Max             =   2
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label lblAtributos 
            BackStyle       =   0  'Transparent
            Caption         =   "Atributos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   100
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Line lineaseparaclases1 
            X1              =   3480
            X2              =   3480
            Y1              =   3840
            Y2              =   2640
         End
         Begin VB.Label lblcaminar 
            BackStyle       =   0  'Transparent
            Caption         =   "Caminar:"
            Height          =   255
            Left            =   2040
            TabIndex        =   74
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label lblcorrer 
            BackStyle       =   0  'Transparent
            Caption         =   "Correr:"
            Height          =   255
            Left            =   2040
            TabIndex        =   73
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lvlvelocidad 
            BackStyle       =   0  'Transparent
            Caption         =   "Velocidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2040
            TabIndex        =   72
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblyspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            Height          =   255
            Left            =   960
            TabIndex        =   68
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label lblxspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   3600
            Width           =   255
         End
         Begin VB.Label lblmapaspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lblspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Spawn"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   64
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label lblvoluntad 
            BackStyle       =   0  'Transparent
            Caption         =   "Voluntad:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   2160
            Width           =   855
         End
         Begin VB.Label lblresistenciaagilidad 
            BackStyle       =   0  'Transparent
            Caption         =   "Resistencia:                   Agilidad:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label lblfuerzainteligencia 
            BackStyle       =   0  'Transparent
            Caption         =   "Fuerza:                          Inteligencia:                             "
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Label lblsprite 
            BackStyle       =   0  'Transparent
            Caption         =   "Sprites Masculino:             Sprites Femenino:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1080
            Width           =   4335
         End
         Begin VB.Label lblnombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "Numero:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jugador"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   44
         Top             =   360
         Width           =   3015
         Begin VB.CheckBox chkbloqpj 
            Height          =   195
            Left            =   2280
            TabIndex        =   45
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Borrar Cuenta al Morir"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   36
         Top             =   3240
         Width           =   5055
         Begin VB.CheckBox chkGUIBars 
            BackColor       =   &H00E0E0E0&
            Caption         =   "GUI Original(NO)"
            Height          =   255
            Left            =   3120
            TabIndex        =   43
            Top             =   360
            Width           =   1815
         End
         Begin VB.CheckBox chkProj 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proyectiles (NO)"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CheckBox chkFS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pantalla Completa (NO)"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   1080
            Width           =   4455
         End
         Begin VB.CheckBox chkDropInvItems 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vaciar Inventario al Morir (Desact)"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   4455
         End
         Begin VB.CheckBox chkFriendSystem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sistema de Amistad (Desact)"
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Feliz"
         Height          =   1215
         Left            =   -71880
         TabIndex        =   34
         Top             =   2040
         Width           =   1935
         Begin VB.CommandButton btnDubExp 
            Caption         =   "  Activar Exp Doble"
            Height          =   615
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtText 
         Height          =   3615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   915
         Width           =   7455
      End
      Begin VB.TextBox txtChat 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   4635
         Width           =   7455
      End
      Begin VB.Frame fraDatabase 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recargar"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   2895
         Begin VB.CommandButton cmdReloadQuests 
            Caption         =   "Misiones"
            Height          =   375
            Left            =   1440
            TabIndex        =   40
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadCombos 
            Caption         =   "Combos"
            Height          =   375
            Left            =   1440
            TabIndex        =   39
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadClasses 
            Caption         =   "Clases"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadMaps 
            Caption         =   "Mapas"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton CmdReloadSpells 
            Caption         =   "Hechizos"
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadShops 
            Caption         =   "Tiendas"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadNPCs 
            Caption         =   "NPCs"
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadItems 
            Caption         =   "Objetos"
            Height          =   375
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadResources 
            Caption         =   "Recursos"
            Height          =   375
            Left            =   1440
            TabIndex        =   22
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdReloadAnimations 
            Caption         =   "Animaciones"
            Height          =   375
            Left            =   1440
            TabIndex        =   21
            Top             =   1200
            Width           =   1215
         End
      End
      Begin VB.Frame fraServer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Servidor"
         Height          =   1575
         Left            =   -71880
         TabIndex        =   16
         Top             =   480
         Width           =   1935
         Begin VB.CommandButton cmdShutDown 
            Caption         =   "Apagar"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "Salir"
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkServerLog 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Logs"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdGSave 
         Caption         =   "Guardar"
         Height          =   255
         Left            =   -69960
         TabIndex        =   15
         Top             =   3195
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Membresía"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   8
         Top             =   1920
         Width           =   6015
         Begin VB.TextBox txtGJoinItem 
            Height          =   285
            Left            =   1080
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtGJoinLvl 
            Height          =   285
            Left            =   3960
            TabIndex        =   10
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtGJoinCost 
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto:"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Req:"
            Height          =   255
            Left            =   2880
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Precio Creacion"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   6015
         Begin VB.TextBox txtGBuyItem 
            Height          =   285
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtGBuyLvl 
            Height          =   285
            Left            =   3960
            TabIndex        =   3
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtGBuyCost 
            Height          =   285
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto:"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Req:"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   840
         End
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7858
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Valor"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Direccion IP"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuenta"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Personaje"
            Object.Width           =   2999
         EndProperty
      End
      Begin VB.Label lblCPS 
         Caption         =   "CPS: 0"
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblCpsLock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "[Desbloq]"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   60
         TabIndex        =   32
         Top             =   600
         Width           =   840
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   495
      Left            =   4800
      TabIndex        =   65
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Menu mnuKick 
      Caption         =   "&Kick"
      Visible         =   0   'False
      Begin VB.Menu mnuKickPlayer 
         Caption         =   "Expulsar"
      End
      Begin VB.Menu mnuDisconnectPlayer 
         Caption         =   "Desconectar"
      End
      Begin VB.Menu mnuBanPlayer 
         Caption         =   "Banear"
      End
      Begin VB.Menu mnuModPlayer 
         Caption         =   "Hacer Mod"
      End
      Begin VB.Menu mnuMapPlayer 
         Caption         =   "Hacer Mapeador"
      End
      Begin VB.Menu mnuDevPlayer 
         Caption         =   "Hacer Desarrollador"
      End
      Begin VB.Menu mnuAdminPlayer 
         Caption         =   "Hacer Admin"
      End
      Begin VB.Menu mnuRemoveAdmin 
         Caption         =   "Quitar Admin"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Player(1).Switches(1) = 1
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub

Private Sub btnDubExp_Click()
    DoubleExp = Not DoubleExp
    If DoubleExp Then
        Call GlobalMsg("Servidor: Exp Doble ACTIVADA.", Green)
        Call TextAdd("Exp Doble ACTIVADA.")
        btnDubExp.Caption = "Desactivar Doble Exp"
    Else
        Call GlobalMsg("Servidor: Exp Doble DESACTIVADA.", Green)
        Call TextAdd("Exp Doble DESACTIVADA.")
        btnDubExp.Caption = "  Activar    Exp Doble"
    End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkDropInvItems_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkDropInvItems.Value Then
        chkDropInvItems.Caption = "Vaciar inventario al morir (Activado)"
    Else
        chkDropInvItems.Caption = "Vaciar inventario al morir (Desactivado)"
    End If
    
    Call PutVar(Path, "OPTIONS", "DropOnDeath", CStr(chkDropInvItems.Value))
End Sub

Private Sub chkFriendSystem_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkFriendSystem.Value Then
        chkFriendSystem.Caption = "Sistema de Amistad (Activado)"
    Else
        chkFriendSystem.Caption = "Sistema de Amistad (Desact)"
    End If
    
    Call PutVar(Path, "OPTIONS", "FriendSystem", CStr(chkFriendSystem.Value))
End Sub

Private Sub chkFS_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkFS.Value Then
        chkFS.Caption = "Pantalla Completa (SI)"
    Else
        chkFS.Caption = "Pantalla Completa (NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "FullScreen", CStr(chkFS.Value))
    Options.FullScreen = chkFS.Value
    
    SendHighIndex
End Sub

Private Sub chkGUIBars_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkGUIBars.Value Then
        chkGUIBars.Caption = "GUI Original(SI)"
    Else
        chkGUIBars.Caption = "GUI Original(NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "OriginalGUIBars", CStr(chkGUIBars.Value))
    Options.OriginalGUIBars = chkGUIBars.Value
    
    SendGUIBarsToAll
End Sub

Private Sub chkProj_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    If chkProj.Value Then
        chkProj.Caption = "Proyectiles (SI)"
    Else
        chkProj.Caption = "Proyectiles (NO)"
    End If
    
    Call PutVar(Path, "OPTIONS", "Projectiles", CStr(chkProj.Value))
    Options.Projectiles = chkProj.Value
End Sub

Private Sub cmdGuardar_Click()

End Sub

Private Sub Cmdbrowser_Click()
frmComunidad.MozillaBrowser1.Navigate ("www.easee.es")
frmComunidad.Visible = True
frmServer.SetFocus
End Sub

Private Sub cmdclima_Click()
Dim clima As Long
Dim intensidad As Long
Dim mapa As Long
Dim niebla As Long
Dim velocidad As Long
Dim opacidad As Long

If Len(txtmapaclima.Text) < 1 Or Len(txtintensidadclima.Text) < 1 Or cmbclima.ListIndex < 0 Or Len(txtniebla.Text) < 1 Or Len(txtvelocidad.Text) < 1 Or Len(txtopacidad.Text) < 1 Then

MsgBox ("Debes completar todos los campos")
Else
clima = cmbclima.ListIndex
intensidad = txtintensidadclima.Text
mapa = txtmapaclima.Text
niebla = txtniebla.Text
opacidad = txtopacidad.Text
velocidad = txtvelocidad.Text
If intensidad > 100 Or niebla > 255 Or opacidad > 255 Or velocidad > 255 Then
MsgBox ("Maximo de Intensidad es 100. Maximo de Niebla, Opacidad y Velocidad es 255.")
Else
Call GlobalMsg("Modificando Mapas Globales", Red)
Call SendClima(mapa, clima, intensidad, niebla, velocidad, opacidad)
End If
End If
End Sub

Private Sub cmdguardarclases_Click()

If txtnombreclase.Text = "" Or txtspritemascclase.Text = "" Or txtspritefemclase.Text = "" Or txtfuerzaclase.Text = "" Or txtresistenciaclase.Text = "" Or txtinteligenciaclase.Text = "" Or txtagilidadclase.Text = "" Or txtvoluntadclase.Text = "" Or txtmapaspawn.Text = "" Or txtxspawn.Text = "" Or txtyspawn.Text = "" Or txtcaminarvelocidad = "" Or txtcorrervelocidad = "" Then
MsgBox ("Debes llenar todos los campos")
Else

PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Name", txtnombreclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "MaleSprite", txtspritemascclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "FemaleSprite", txtspritefemclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Strength", txtfuerzaclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Endurance", txtresistenciaclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Intelligence", txtinteligenciaclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Agility", txtagilidadclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "WillPower", txtvoluntadclase.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Mapa", txtmapaspawn.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "X", txtxspawn.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Y", txtyspawn.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "VCaminar", txtcaminarvelocidad.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "VCorrer", txtcorrervelocidad.Text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Visible", chkvisible.Value

End If
End Sub



Private Sub cmdReloadCombos_Click()
Dim I As Long
    Call LoadCombos
    Call TextAdd("Combos Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendCombos I
        End If
    Next
End Sub

Private Sub cmdReloadQuests_Click()
Dim I As Long
    Call LoadQuests
    Call TextAdd("Misiones Actualizadas.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendQuests I
        End If
    Next
End Sub
Private Sub cmdGSave_Click()
    Options.Buy_Cost = frmServer.txtGBuyCost.Text
    Options.Buy_Lvl = frmServer.txtGBuyLvl.Text
    Options.Buy_Item = frmServer.txtGBuyItem.Text
    Options.Join_Cost = frmServer.txtGJoinCost.Text
    Options.Join_Lvl = frmServer.txtGJoinLvl.Text
    Options.Join_Item = frmServer.txtGJoinItem.Text
    SaveOptions
End Sub



Private Sub Combo1_Change()

End Sub



Private Sub Frame5_DragDrop(Source As Control, x As Single, y As Single)

End Sub


Private Sub lblCPSLock_Click()
    If CPSUnlock Then
        CPSUnlock = False
        lblCpsLock.Caption = "[Bloquear]"
    Else
        CPSUnlock = True
        lblCpsLock.Caption = "[Desbloq]"
    End If
End Sub


Private Sub sldclasenum_Change() 'Cortesia de EaSee Engine (que lindo)
Dim filename As String
filename = App.Path & "\data\classes.ini"
lblnumero.Caption = "Numero: " & (sldclasenum.Value)
frmServer.txtnombreclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Name")
        frmServer.txtspritemascclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "MaleSprite")
        frmServer.txtspritefemclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "FemaleSprite")
        frmServer.txtfuerzaclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Strength")
        frmServer.txtresistenciaclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Endurance")
        frmServer.txtinteligenciaclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Intelligence")
        frmServer.txtagilidadclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Agility")
        frmServer.txtvoluntadclase.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "WillPower")
        frmServer.txtmapaspawn.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Mapa")
        frmServer.txtxspawn.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "X")
        frmServer.txtyspawn.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Y")
        frmServer.txtcaminarvelocidad.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "VCaminar")
        frmServer.txtcorrervelocidad.Text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "VCorrer")
        frmServer.chkvisible.Value = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Visible")
End Sub

' ********************
' ** Winsock object **
' ********************
Private Sub Socket_ConnectionRequest(index As Integer, ByVal requestID As Long)
    Call AcceptConnection(index, requestID)
End Sub

Private Sub Socket_Accept(index As Integer, SocketId As Integer)
    Call AcceptConnection(index, SocketId)
End Sub

Private Sub Socket_DataArrival(index As Integer, ByVal bytesTotal As Long)

    If IsConnected(index) Then
        Call IncomingData(index, bytesTotal)
    End If

End Sub

Private Sub Socket_Close(index As Integer)
    Call CloseSocket(index)
End Sub

' ********************
Private Sub chkServerLog_Click()

    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If

End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim I As Long
    Call LoadClasses
    Call TextAdd("Clases Actualizadas.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendClasses I
            
        End If
    Next
End Sub

Private Sub cmdReloadItems_Click()
Dim I As Long
    Call LoadItems
    Call TextAdd("Objetos Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendItems I
        End If
    Next
End Sub

Private Sub cmdReloadMaps_Click()
Dim I As Long
    Call LoadMaps
    Call TextAdd("Mapas Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            PlayerWarp I, GetPlayerMap(I), GetPlayerX(I), GetPlayerY(I)
        End If
    Next
End Sub

Private Sub cmdReloadNPCs_Click()
Dim I As Long
    Call LoadNpcs
    Call TextAdd("NPCs Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendNpcs I
        End If
    Next
End Sub

Private Sub cmdReloadShops_Click()
Dim I As Long
    Call LoadShops
    Call TextAdd("Tiendas Actualizadas.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendShops I
        End If
    Next
End Sub

Private Sub cmdReloadSpells_Click()
Dim I As Long
    Call LoadSpells
    Call TextAdd("Hechizos Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendSpells I
        End If
    Next
End Sub

Private Sub cmdReloadResources_Click()
Dim I As Long
    Call LoadResources
    Call TextAdd("Recursos Actualizados.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendResources I
        End If
    Next
End Sub

Private Sub cmdReloadAnimations_Click()
Dim I As Long
    Call LoadAnimations
    Call TextAdd("Animaciones Actualizadas.")
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendAnimations I
        End If
    Next
End Sub

Private Sub cmdShutDown_Click()
    If isShuttingDown Then
        isShuttingDown = False
        cmdShutDown.Caption = "Apagar"
        GlobalMsg "Apagado Cancelado.", BrightBlue
    Else
        isShuttingDown = True
        cmdShutDown.Caption = "Cancelar"
    End If
End Sub

Private Sub Form_Load()
    Call SetData
End Sub

Private Sub Form_Resize()

    If frmServer.WindowState = vbMinimized Then
        frmServer.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Call DestroyServer
End Sub

Private Sub lvwInfo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'When a ColumnHeader object is clicked, the ListView control is sorted by the subitems of that column.
    'Set the SortKey to the Index of the ColumnHeader - 1
    'Set Sorted to True to sort the list.
    If lvwInfo.SortOrder = lvwAscending Then
        lvwInfo.SortOrder = lvwDescending
    Else
        lvwInfo.SortOrder = lvwAscending
    End If

    lvwInfo.SortKey = ColumnHeader.index - 1
    lvwInfo.Sorted = True
End Sub

Private Sub Text2_Change()

End Sub

Private Sub tmrGetTime_Timer()
Dim Display_S As String, Display_M As String, Display_H As String
    'get the time variables
    
    If Time_Seconds + 1 = 60 And Time_Minutes + 1 = 60 And Time_Hours + 1 = 24 Then
    Time_Hours = 0
    Time_Seconds = 0
    Time_Minutes = 0
    End If
        
    If Time_Seconds + 1 = 60 Then
        If Time_Minutes + 1 = 60 Then
            Time_Hours = Time_Hours + 1
            Time_Minutes = 0
        Else
            Time_Minutes = Time_Minutes + 1
            Time_Seconds = 0
        End If
    Else
        Time_Seconds = Time_Seconds + 1
    End If
    
    
    'prepare them
    Display_S = Time_Seconds
    Display_M = Time_Minutes
    Display_H = Time_Hours
    If Time_Seconds < 10 Then Display_S = "0" & Time_Seconds
    If Time_Minutes < 10 Then Display_M = "0" & Time_Minutes
    If Time_Hours < 10 Then Display_H = "0" & Time_Hours
    
    lbltiempo.Caption = Display_H & ":" & Display_M & ":" & Display_S
    
End Sub




Private Sub txtagilidadclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub





Private Sub txtcaminarvelocidad_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtcorrervelocidad_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtfuerzaclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtguardarmaxmin_Click()

Dim clasesexistentes As Byte
clasesexistentes = Max_Classes
Dim clasesdeseadas As Byte
Dim filename As String
Dim I As Byte
Dim x As Long

filename = App.Path & "\data\classes.ini"


If CInt(txtmaxclases.Text) < 2 Then 'Guardado de Max_Classes
MsgBox ("Clases Debe ser Igual o Mayor a 2")
Else
clasesdeseadas = txtmaxclases.Text
PutVar App.Path & "\data\classes.ini", "INIT", "MaxClasses", txtmaxclases.Text

If clasesexistentes < clasesdeseadas Then 'Si pusiste mas Clases lo detecta y crea datos del fichero
MsgBox ("Existentes" & clasesexistentes)
MsgBox ("Nuevas" & clasesdeseadas)
  For I = clasesexistentes + 1 To clasesdeseadas

       Call PutVar(filename, "CLASS" & I, "Name", "NuevaClase")
        Call PutVar(filename, "CLASS" & I, "Malesprite", "1")
        Call PutVar(filename, "CLASS" & I, "Femalesprite", "1")
        Call PutVar(filename, "CLASS" & I, "Strength", "1")
        Call PutVar(filename, "CLASS" & I, "Endurance", "1")
        Call PutVar(filename, "CLASS" & I, "Intelligence", "1")
        Call PutVar(filename, "CLASS" & I, "Agility", "1")
        Call PutVar(filename, "CLASS" & I, "Willpower", "1")
        Call PutVar(filename, "CLASS" & I, "Mapa", "1")
        Call PutVar(filename, "CLASS" & I, "X", "1")
        Call PutVar(filename, "CLASS" & I, "Y", "1")
        Call PutVar(filename, "CLASS" & I, "VCaminar", "10")
        Call PutVar(filename, "CLASS" & I, "VCorrer", "16")
        Call PutVar(filename, "CLASS" & I, "Visible", "0")

        ' loop for items & values
        For x = 1 To 1 'Item Inicial 1
            Call PutVar(filename, "CLASS" & I, "StartItem" & x, "1")
            Call PutVar(filename, "CLASS" & I, "StartValue" & x, "1")
        Next
        ' loop for spells
        For x = 1 To 1 'Hechizo Inicial 1
            Call PutVar(filename, "CLASS" & I, "StartSpell" & x, "1")
        Next
    Next
    Max_Classes = txtmaxclases.Text
    MsgBox ("Guardado")
    Else
    MsgBox ("Guardado")
End If
End If
End Sub

Private Sub txtinteligenciaclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtintensidadclima_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txtmapaclima_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtmapaspawn_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtniebla_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtopacidad_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtresistenciaclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtspritefemclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 44 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtspritemascclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Or KeyAscii = 44 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtText_GotFocus()
    txtChat.SetFocus
End Sub

Private Sub txtChat_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If LenB(Trim$(txtChat.Text)) > 0 Then
            Call GlobalMsg(txtChat.Text, White)
            Call TextAdd("Servidor: " & txtChat.Text)
            txtChat.Text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (I)

        If I < 10 Then
            frmServer.lvwInfo.ListItems(I).Text = "00" & I
        ElseIf I < 100 Then
            frmServer.lvwInfo.ListItems(I).Text = "0" & I
        Else
            frmServer.lvwInfo.ListItems(I).Text = I
        End If

        frmServer.lvwInfo.ListItems(I).SubItems(1) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(2) = vbNullString
        frmServer.lvwInfo.ListItems(I).SubItems(3) = vbNullString
    Next

End Sub

Private Sub lvwInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuKick
    End If

End Sub

Private Sub mnuKickPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call AlertMsg(FindPlayer(Name), "Has sido expulsado del Servidor!")
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Privilegios de Administrador Activados.", BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Privilegios de Administrador Desactivados.", BrightRed)
    End If

End Sub

Sub mnuModPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 1)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Privilegios de Moderador Activados.", BrightCyan)
    End If

End Sub

Sub mnuMapPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 2)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Privilegios de Mapeador Activados.", BrightCyan)
    End If

End Sub

Sub mnuDevPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de Linea" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 3)
        Call SendPlayerData(FindPlayer(Name))
        Call PlayerMsg(FindPlayer(Name), "Privilegios de Desarrollador Activados.", BrightCyan)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.Text)
    End Select

End Sub



Private Sub txtvelocidad_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtvoluntadclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtwalk_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtrun_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtxspawn_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub



Private Sub txtyspawn_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub
