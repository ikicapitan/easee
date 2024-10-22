VERSION 5.00
Begin VB.Form frmEditor_Map 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Mapa"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14655
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   497
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   977
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CheckBox chckgrilla 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grilla"
      Height          =   180
      Left            =   5760
      TabIndex        =   100
      ToolTipText     =   "Permite ver las grillas para poder trabajar sobre el Mapa de forma mas comoda."
      Top             =   5280
      Width           =   1455
   End
   Begin VB.PictureBox picAttributes 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   7320
      ScaleHeight     =   7215
      ScaleWidth      =   7095
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Frame fraSoundEffect 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Efecto de Sonido"
         Height          =   1455
         Left            =   120
         TabIndex        =   89
         Top             =   5640
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdSoundEffectOk 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   91
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cmbSoundEffect 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3332
            Left            =   240
            List            =   "frmEditor_Map.frx":3342
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraSlide 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deslizar"
         Height          =   1455
         Left            =   120
         TabIndex        =   85
         Top             =   4080
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbSlide 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":335D
            Left            =   240
            List            =   "frmEditor_Map.frx":336D
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdSlide 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   86
            Top             =   840
            Width           =   1455
         End
      End
      Begin VB.Frame fraTrap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trampa"
         Height          =   1575
         Left            =   120
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
         Begin VB.HScrollBar scrlTrap 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   83
            Top             =   600
            Width           =   2895
         End
         Begin VB.CommandButton cmdTrap 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   82
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblTrap 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraHeal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Curar"
         Height          =   1815
         Left            =   3240
         TabIndex        =   76
         Top             =   3600
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ComboBox cmbHeal 
            Height          =   300
            ItemData        =   "frmEditor_Map.frx":3388
            Left            =   240
            List            =   "frmEditor_Map.frx":3392
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   240
            Width           =   2895
         End
         Begin VB.CommandButton cmdHeal 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   78
            Top             =   1200
            Width           =   1455
         End
         Begin VB.HScrollBar scrlHeal 
            Height          =   255
            Left            =   240
            Max             =   10000
            TabIndex        =   77
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label lblHeal 
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad: 0"
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.Frame fraNpcSpawn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Npc Spawn"
         Height          =   2655
         Left            =   3000
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox lstNpc 
            Height          =   780
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   2895
         End
         Begin VB.HScrollBar scrlNpcDir 
            Height          =   255
            Left            =   240
            Max             =   3
            TabIndex        =   37
            Top             =   1560
            Width           =   2895
         End
         Begin VB.CommandButton cmdNpcSpawn 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   36
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblNpcDir 
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion: Arriba"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   1320
            Width           =   2535
         End
      End
      Begin VB.Frame fraResource 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto"
         Height          =   1695
         Left            =   2280
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdResourceOk 
            Caption         =   "Ok"
            Height          =   375
            Left            =   960
            TabIndex        =   32
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   240
            Max             =   100
            Min             =   1
            TabIndex        =   31
            Top             =   600
            Value           =   1
            Width           =   2895
         End
         Begin VB.Label lblResource 
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame fraMapWarp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mapa Transportaci�n"
         Height          =   2775
         Left            =   2760
         TabIndex        =   58
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapWarp 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   1080
            TabIndex        =   65
            Top             =   2160
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapWarpY 
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarpX 
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMapWarp 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   60
            Top             =   480
            Value           =   1
            Width           =   3135
         End
         Begin VB.Label lblMapWarpY 
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblMapWarpX 
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblMapWarp 
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraShop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tienda"
         Height          =   1335
         Left            =   3360
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdShop 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   960
            TabIndex        =   68
            Top             =   720
            Width           =   1215
         End
         Begin VB.ComboBox cmbShop 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame fraKeyOpen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir Llave"
         Height          =   2055
         Left            =   2880
         TabIndex        =   52
         Top             =   2040
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdKeyOpen 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   1080
            TabIndex        =   57
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlKeyY 
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1080
            Width           =   3015
         End
         Begin VB.HScrollBar scrlKeyX 
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblKeyY 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label lblKeyX 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame fraMapKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Llave Mapa"
         Height          =   1815
         Left            =   360
         TabIndex        =   46
         Top             =   2400
         Visible         =   0   'False
         Width           =   3375
         Begin VB.PictureBox picMapKey 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   51
            Top             =   600
            Width           =   480
         End
         Begin VB.CommandButton cmdMapKey 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   1080
            TabIndex        =   50
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CheckBox chkMapKey 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Debes tener la llave para usar."
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   960
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.HScrollBar scrlMapKey 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   48
            Top             =   600
            Value           =   1
            Width           =   2535
         End
         Begin VB.Label lblMapKey 
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto: Ninguno"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraMapItem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto Mapa"
         Height          =   1815
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CommandButton cmdMapItem 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   1200
            TabIndex        =   45
            Top             =   1200
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapItemValue 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   44
            Top             =   840
            Value           =   1
            Width           =   2535
         End
         Begin VB.HScrollBar scrlMapItem 
            Height          =   255
            Left            =   120
            Max             =   10
            Min             =   1
            TabIndex        =   43
            Top             =   480
            Value           =   1
            Width           =   2535
         End
         Begin VB.PictureBox picMapItem 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2760
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   42
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lblMapItem 
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto: Ninguno x0"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Propiedades"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Propiedades Avanzadas del Mapa."
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modo de Edici�n"
      Height          =   1335
      Left            =   5760
      TabIndex        =   24
      ToolTipText     =   "Aqui podras seleccionar la herramienta con la que desees trabajar sobre el Mapa."
      Top             =   5520
      Width           =   1455
      Begin VB.OptionButton optEvent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   " Eventos"
         Height          =   255
         Left            =   240
         TabIndex        =   88
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optBlock 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "  Bloqueo"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optAttribs 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "  Atributos"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optLayers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Capas"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.HScrollBar scrlPictureX 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   5295
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   120
      ScaleHeight     =   352
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   352
      TabIndex        =   14
      Top             =   120
      Width           =   5280
   End
   Begin VB.VScrollBar scrlPictureY 
      Height          =   5295
      Left            =   5400
      Max             =   255
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame fraTileSet 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tileset: 0"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Tileset o Galeria de Imagenes para que Edites Tus Mapas."
      Top             =   6000
      Width           =   5535
      Begin VB.HScrollBar scrlTileSet 
         Height          =   255
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   1
         Top             =   360
         Value           =   1
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame fraLayers 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Capas"
      Height          =   5055
      Left            =   5760
      TabIndex        =   15
      ToolTipText     =   $"frmEditor_Map.frx":33A4
      Top             =   120
      Width           =   1455
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tile al Azar"
         Height          =   1215
         Left            =   0
         TabIndex        =   95
         Top             =   2760
         Width           =   1455
         Begin VB.CommandButton cmdRandomTile 
            Caption         =   "Colocar"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   840
            Width           =   1215
         End
         Begin VB.HScrollBar scrlFrequency 
            Height          =   255
            Left            =   120
            Max             =   100
            Min             =   1
            TabIndex        =   96
            Top             =   480
            Value           =   75
            Width           =   1215
         End
         Begin VB.Label lblFrequency 
            BackStyle       =   0  'Transparent
            Caption         =   "Frec.: 75"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.HScrollBar scrlAutotile 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   93
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "Llenar"
         Height          =   390
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Rellena el Mapa Automaticamente con el Tile Seleccionado en la Capa Seleccionada."
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Superior1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "M�scara1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Suelo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Superior2"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optLayer 
         BackColor       =   &H00E0E0E0&
         Caption         =   "M�scara2"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Vacia la Capa Seleccionada."
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblAutotile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Normal"
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.Frame fraAttribs 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Atributos"
      Height          =   5055
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optPlayerSpawn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Punto Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sonido"
         Height          =   270
         Left            =   120
         TabIndex        =   92
         Top             =   3600
         Width           =   1215
      End
      Begin VB.OptionButton optSlide 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Deslizar"
         Height          =   270
         Left            =   120
         TabIndex        =   75
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton optTrap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Trampa"
         Height          =   270
         Left            =   120
         TabIndex        =   74
         Top             =   3120
         Width           =   1215
      End
      Begin VB.OptionButton optHeal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Curar"
         Height          =   270
         Left            =   120
         TabIndex        =   73
         Top             =   2880
         Width           =   1215
      End
      Begin VB.OptionButton optBank 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
         Height          =   270
         Left            =   120
         TabIndex        =   72
         Top             =   2640
         Width           =   1215
      End
      Begin VB.OptionButton optShop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tienda"
         Height          =   270
         Left            =   120
         TabIndex        =   69
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton optNpcSpawn 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Npc Spawn"
         Height          =   270
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton optDoor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Puerta"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optResource 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recurso"
         Height          =   240
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1215
      End
      Begin VB.OptionButton optKeyOpen 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir Llave"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBlocked 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bloqueado"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optWarp 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transportar"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "Clear"
         Height          =   390
         Left            =   120
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
      End
      Begin VB.OptionButton optItem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto"
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optNpcAvoid 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Npc Evitar"
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optKey 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Llave"
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arrastra el rat�n para seleccionar m�s de 1 Tile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   70
      Top             =   5760
      Width           =   5535
   End
End
Attribute VB_Name = "frmEditor_Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorHealType = cmbHeal.ListIndex + 1
    MapEditorHealAmount = scrlHeal.Value
    picAttributes.Visible = False
    fraHeal.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdKeyOpen_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    KeyOpenEditorX = scrlKeyX.Value
    KeyOpenEditorY = scrlKeyY.Value
    picAttributes.Visible = False
    fraKeyOpen.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ItemEditorNum = scrlMapItem.Value
    ItemEditorValue = scrlMapItemValue.Value
    picAttributes.Visible = False
    fraMapItem.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdMapItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapKey_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    KeyEditorNum = scrlMapKey.Value
    KeyEditorTake = chkMapKey.Value
    picAttributes.Visible = False
    fraMapKey.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdMapKey_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdMapWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorWarpMap = scrlMapWarp.Value
    EditorWarpX = scrlMapWarpX.Value
    EditorWarpY = scrlMapWarpY.Value
    picAttributes.Visible = False
    fraMapWarp.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdMapWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdNpcSpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpawnNpcNum = lstNpc.ListIndex + 1
    SpawnNpcDir = scrlNpcDir.Value
    picAttributes.Visible = False
    fraNpcSpawn.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdResourceOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResourceEditorNum = scrlResource.Value
    picAttributes.Visible = False
    fraResource.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdResourceOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorShop = cmbShop.ListIndex
    picAttributes.Visible = False
    fraShop.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorSlideDir = cmbSlide.ListIndex
    picAttributes.Visible = False
    fraSlide.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSoundEffectOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorSound = soundCache(cmbSoundEffect.ListIndex + 1)
    picAttributes.Visible = False
    fraSoundEffect.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSoundEffectOk_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorHealAmount = scrlTrap.Value
    picAttributes.Visible = False
    fraTrap.Visible = False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MapEditor", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L1")
fraLayers.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L2")
optLayer(1).Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L3")
optLayer(2).Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L4")
optLayer(3).Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L5")
optLayer(4).Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L6")
optLayer(5).Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L8")
chckgrilla.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L9")
Frame2.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L10")
optLayers.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L12")
optAttribs.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L13")
optBlock.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L14")
optEvent.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L15")
Label1.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L16")
fraSoundEffect.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L17")
fraSlide.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L18")
fraTrap.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L19")
lblTrap.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L20")
fraHeal.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L21")
lblHeal.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L22")
lblNpcDir.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L23")
fraResource.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L24")
lblResource.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L25")
fraMapWarp.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L26")
lblMapWarp.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L27")
fraKeyOpen.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L28")
fraShop.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L29")
fraMapKey.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L30")
lblMapKey.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L31")
chkMapKey.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L32")
fraMapItem.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L33")
lblMapItem.Caption = trad
trad = GetVar(App.Path & Lang, "MapEditor", "L34")
Frame3.Caption = trad


Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' move the entire attributes box on screen
    picAttributes.Left = 8
    picAttributes.Top = 8
    
    GraphicSelX = 0
    GraphicSelY = 0
    GraphicSelX2 = 0
    GraphicSelY2 = 0
    
    PopulateLists
    
    cmbSoundEffect.Clear
    For I = 1 To UBound(soundCache)
        cmbSoundEffect.AddItem (soundCache(I))
    Next
    cmbSoundEffect.ListIndex = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optDoor_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optDoor_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optHeal_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraHeal.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optHeal_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optLayers_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If optLayers.Value Then
        fraLayers.Visible = True
        fraAttribs.Visible = False
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optLayers_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optAttribs_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If optAttribs.Value Then
        fraLayers.Visible = False
        fraAttribs.Visible = True
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optAttribs_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optNpcSpawn_Click()
Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstNpc.Clear
    
    For n = 1 To MAX_MAP_NPCS
        If Map.NPC(n) > 0 Then
            lstNpc.AddItem n & ": " & NPC(Map.NPC(n)).name
        Else
            lstNpc.AddItem n & ": No Npc"
        End If
    Next n
    
    scrlNpcDir.Value = 0
    lstNpc.ListIndex = 0
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraNpcSpawn.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optNpcSpawn_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraResource.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optResource_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraShop.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optShop_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSlide_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSlide.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optSlide_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraSoundEffect.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optSound_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optTrap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraTrap.Visible = True
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optTrap_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSend_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmEditor_Events.Visible Then
        If Msgbox("El Editor de Eventos esta abierto. Continuar enviando datos al mapa no guardara el evento en edicion. Continuar?", vbYesNo) = vbYes Then
            Call MapEditorSend
        End If
    Else
        Call MapEditorSend
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSend_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Events.Visible = False
    Call MapEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdProperties_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Load frmEditor_MapProperties
    MapEditorProperties
    frmEditor_MapProperties.Show vbModal
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdProperties_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optWarp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapWarp.Visible = True
    
    scrlMapWarp.max = MAX_MAPS
    scrlMapWarp.Value = 1
    scrlMapWarpX.max = MAX_BYTE
    scrlMapWarpY.max = MAX_BYTE
    scrlMapWarpX.Value = 0
    scrlMapWarpY.Value = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optWarp_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapItem.Visible = True

    scrlMapItem.max = MAX_ITEMS
    scrlMapItem.Value = 1
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optItem_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optKey_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    picAttributes.Visible = True
    fraMapKey.Visible = True
    
    scrlMapKey.max = MAX_ITEMS
    scrlMapKey.Value = 1
    chkMapKey.Value = 1
    lblMapKey.Caption = "Objeto: " & Trim$(Item(scrlMapKey.Value).name)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optKey_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optKeyOpen_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ClearAttributeDialogue
    fraKeyOpen.Visible = True
    picAttributes.Visible = True
    
    scrlKeyX.max = Map.MaxX
    scrlKeyY.max = Map.MaxY
    scrlKeyX.Value = 0
    scrlKeyY.Value = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optKeyOpen_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdFill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapEditorFillLayer
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdFill_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearLayer
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdClear_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdClear2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorClearAttribs
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdClear2_Click", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorChooseTile(Button, X, Y)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picBack_MouseDown", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    X = X + (frmEditor_Map.scrlPictureX.Value * 32)
    Y = Y + (frmEditor_Map.scrlPictureY.Value * 32)
    Call MapEditorDrag(Button, X, Y)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picBack_MouseMove", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAutotile_Change()
    Select Case scrlAutotile.Value
        Case 0 ' normal
            lblAutotile.Caption = "Normal"
        Case 1 ' autotile
            lblAutotile.Caption = "Autotile (VX)"
        Case 2 ' fake autotile
            lblAutotile.Caption = "Fake (VX)"
        Case 3 ' animated
            lblAutotile.Caption = "Animated (VX)"
        Case 4 ' cliff
            lblAutotile.Caption = "Cliff (VX)"
        Case 5 ' waterfall
            lblAutotile.Caption = "Waterfall (VX)"
        Case 6 ' autotile
            lblAutotile.Caption = "Autotile (XP)"
        Case 7 ' fake autotile
            lblAutotile.Caption = "Fake (XP)"
        Case 8 ' animated
            lblAutotile.Caption = "Animated (XP)"
        Case 9 ' cliff
            lblAutotile.Caption = "Cliff (XP)"
        Case 10 ' waterfall
            lblAutotile.Caption = "Waterfall (XP)"
    End Select
End Sub

Private Sub scrlFrequency_Change()
    lblFrequency.Caption = "Frec.: " & scrlFrequency.Value
End Sub

Private Sub scrlHeal_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L21")
lblHeal.Caption = trad
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblHeal.Caption = "Cantidad: " & scrlHeal.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlHeal_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblKeyX.Caption = "X: " & scrlKeyX.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlKeyX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlKeyX_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlKeyX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblKeyY.Caption = "Y: " & scrlKeyY.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlKeyY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlKeyY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlKeyY_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlKeyY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTrap_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L19")
lblTrap.Caption = trad
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblTrap.Caption = trad & " " & scrlTrap.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlTrap_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Change()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
        
    If Item(scrlMapItem.Value).Type = ITEM_TYPE_CURRENCY Or Item(scrlMapItem.Value).Stackable > 0 Then
        scrlMapItemValue.Enabled = True
    Else
        scrlMapItemValue.Value = 1
        scrlMapItemValue.Enabled = False
    End If

    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapItem_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItem_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItem_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapItem_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapItem.Caption = Trim$(Item(scrlMapItem.Value).name) & " x" & scrlMapItemValue.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapItemValue_change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapItemValue_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapItemValue_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapItemValue_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapKey_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L30")
lblMapKey.Caption = trad
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapKey.Caption = trad & " " & Trim$(Item(scrlMapKey.Value).name)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapKey_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapKey_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapKey_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapKey_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L26")
lblMapWarp.Caption = trad
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarp.Caption = trad & " " & scrlMapWarp.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarp_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarp_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarp_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarp_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpX.Caption = "X: " & scrlMapWarpX.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarpX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpX_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarpX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblMapWarpY.Caption = "Y: " & scrlMapWarpY.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarpY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMapWarpY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlMapWarpY_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMapWarpY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L22")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case scrlNpcDir.Value
        Case DIR_DOWN
            lblNpcDir = trad & " Down"
        Case DIR_UP
            lblNpcDir = trad & " Up"
        Case DIR_LEFT
            lblNpcDir = trad & " Left"
        Case DIR_RIGHT
            lblNpcDir = trad & " Right"
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlNpcDir_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNpcDir_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlNpcDir_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlNpcDir_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "MapEditor", "L24")
lblResource.Caption = trad
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblResource.Caption = trad & " " & Resource(scrlResource.Value).name
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlResource_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlResource_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlResource_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlResource_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorTileScroll
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPictureX_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MapEditorTileScroll
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPictureY_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureX_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPictureY_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPictureX_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPictureY_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPictureY_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPictureY_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fraTileSet.Caption = "Tileset: " & scrlTileSet.Value
    
    frmEditor_Map.scrlPictureY.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Height \ PIC_Y) - (frmEditor_Map.picBack.Height \ PIC_Y)
    frmEditor_Map.scrlPictureX.max = (Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width \ PIC_X) - (frmEditor_Map.picBack.Width \ PIC_X)
    
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlTileSet_Change", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTileSet_Scroll()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlTileSet_Change
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlTileSet_Scroll", "frmEditor_Map", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
