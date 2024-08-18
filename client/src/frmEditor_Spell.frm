VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Hechizos"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13800
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
   Icon            =   "frmEditor_Spell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8640
      TabIndex        =   61
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar y Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7560
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   12240
      TabIndex        =   5
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Propiedades de Hechizo"
      Height          =   7335
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos Adicionales"
         Height          =   6975
         Left            =   6840
         TabIndex        =   62
         Top             =   240
         Width           =   3375
         Begin VB.TextBox txtVeneno 
            Height          =   270
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   101
            Text            =   "0"
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox txtmapa 
            Height          =   270
            Left            =   2640
            MaxLength       =   3
            TabIndex        =   97
            Text            =   "0"
            Top             =   6480
            Width           =   495
         End
         Begin VB.TextBox txty 
            Height          =   270
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   96
            Text            =   "0"
            Top             =   6480
            Width           =   495
         End
         Begin VB.TextBox txtx 
            Height          =   270
            Left            =   1440
            MaxLength       =   3
            TabIndex        =   95
            Text            =   "0"
            Top             =   6480
            Width           =   495
         End
         Begin VB.CheckBox chktransportar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Transportar"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            ToolTipText     =   "Transporta al Objetivo a las Coordenadas Indicadas."
            Top             =   6480
            Width           =   1215
         End
         Begin VB.TextBox txtsprite 
            Height          =   270
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   93
            Text            =   "0"
            Top             =   6120
            Width           =   495
         End
         Begin VB.CheckBox chksprite 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sprite"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "Modifica el Sprite del Objetivo."
            Top             =   6120
            Width           =   855
         End
         Begin VB.TextBox txtarco 
            Height          =   270
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   90
            Text            =   "0"
            Top             =   5760
            Width           =   495
         End
         Begin VB.CheckBox chkarco 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Arco"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   5760
            Width           =   735
         End
         Begin VB.CheckBox chkbuff 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Buff Stat (Destildado para DeBuff)"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            ToolTipText     =   "Afecta los atributos del objetivo arriba especificados."
            Top             =   5400
            Width           =   2895
         End
         Begin VB.TextBox txtvol 
            Height          =   270
            Left            =   2280
            MaxLength       =   3
            TabIndex        =   86
            Text            =   "0"
            Top             =   4660
            Width           =   615
         End
         Begin VB.TextBox txtint 
            Height          =   270
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   85
            Text            =   "0"
            Top             =   4290
            Width           =   615
         End
         Begin VB.TextBox txtagi 
            Height          =   270
            Left            =   840
            MaxLength       =   3
            TabIndex        =   84
            Text            =   "0"
            Top             =   5000
            Width           =   615
         End
         Begin VB.TextBox txtdes 
            Height          =   270
            Left            =   860
            MaxLength       =   3
            TabIndex        =   83
            Text            =   "0"
            Top             =   4640
            Width           =   615
         End
         Begin VB.TextBox txtfza 
            Height          =   270
            Left            =   720
            MaxLength       =   3
            TabIndex        =   82
            Text            =   "0"
            Top             =   4300
            Width           =   615
         End
         Begin VB.TextBox txtcorrer 
            Height          =   270
            Left            =   840
            MaxLength       =   3
            TabIndex        =   75
            Text            =   "0"
            Top             =   3840
            Width           =   495
         End
         Begin VB.TextBox txtcaminar 
            Height          =   270
            Left            =   120
            MaxLength       =   3
            TabIndex        =   74
            Text            =   "0"
            Top             =   3840
            Width           =   495
         End
         Begin VB.CheckBox chkVelocidad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Velocidad (Destildado para Restar)"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            ToolTipText     =   "Aumenta o Quita Velocidad al Objetivo."
            Top             =   3480
            Width           =   3015
         End
         Begin VB.CheckBox chkveneno 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Veneno"
            Height          =   255
            Left            =   1680
            TabIndex        =   72
            ToolTipText     =   "Envenena al Objetivo."
            Top             =   3120
            Width           =   1335
         End
         Begin VB.CheckBox chkinvisibilidad 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Invisibilidad"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            ToolTipText     =   "Hace Invisible al Objetivo."
            Top             =   3120
            Width           =   1335
         End
         Begin VB.CheckBox chkconfusion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Confusion"
            Height          =   255
            Left            =   1680
            TabIndex        =   70
            ToolTipText     =   "Causa confusion al Objetivo."
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CheckBox chkparalizar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Paralizar"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            ToolTipText     =   "Paraliza al Objetivo."
            Top             =   2760
            Width           =   1215
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   65
            Top             =   2040
            Width           =   3135
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   64
            Top             =   600
            Width           =   3135
         End
         Begin VB.HScrollBar scrlElement 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   63
            Top             =   1320
            Width           =   3135
         End
         Begin VB.Label lblMapa 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   " Mapa"
            Height          =   255
            Left            =   2640
            TabIndex        =   100
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label lblYY 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "    y"
            Height          =   255
            Left            =   2040
            TabIndex        =   99
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label lblXX 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "    x"
            Height          =   255
            Left            =   1440
            TabIndex        =   98
            Top             =   6720
            Width           =   615
         End
         Begin VB.Label lblsprite 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sprite:"
            Height          =   255
            Left            =   1200
            TabIndex        =   92
            Top             =   6170
            Width           =   495
         End
         Begin VB.Label lblarco 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Item:"
            Height          =   255
            Left            =   1200
            TabIndex        =   89
            Top             =   5790
            Width           =   495
         End
         Begin VB.Label lblcamcorr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "(Caminar/Correr)"
            Height          =   255
            Left            =   1440
            TabIndex        =   81
            Top             =   3885
            Width           =   1335
         End
         Begin VB.Label lblVol 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Voluntad:"
            Height          =   255
            Left            =   1560
            TabIndex        =   80
            ToolTipText     =   "Buff de Voluntad"
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label lblInt 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Inteligencia:"
            Height          =   255
            Left            =   1560
            TabIndex        =   79
            ToolTipText     =   "Buff de Inteligencia"
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label lblagi 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agilidad:"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            ToolTipText     =   "Buff de Agilidad"
            Top             =   5040
            Width           =   735
         End
         Begin VB.Label lbldes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Destreza:"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            ToolTipText     =   "Buff de Destreza"
            Top             =   4680
            Width           =   735
         End
         Begin VB.Label lblfza 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Fuerza:"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            ToolTipText     =   "Buff de Fuerza"
            Top             =   4320
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00404040&
            BorderStyle     =   6  'Inside Solid
            DrawMode        =   14  'Copy Pen
            X1              =   120
            X2              =   3240
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Golpe Neutral: 0"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            UseMnemonic     =   0   'False
            Width           =   1245
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Golpe Luz: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   67
            Top             =   360
            UseMnemonic     =   0   'False
            Width           =   945
         End
         Begin VB.Label lblElement 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Golpe Oscuridad: 0"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   66
            Top             =   1080
            UseMnemonic     =   0   'False
            Width           =   1470
         End
      End
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   1440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   57
         Top             =   6240
         Width           =   2775
      End
      Begin VB.HScrollBar scrlCombatLvl 
         Height          =   255
         LargeChange     =   10
         Left            =   5760
         Max             =   100
         TabIndex        =   56
         Top             =   6840
         Width           =   975
      End
      Begin VB.ComboBox cmdCombatType 
         Height          =   300
         ItemData        =   "frmEditor_Spell.frx":08CA
         Left            =   4800
         List            =   "frmEditor_Spell.frx":08DA
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   6360
         Width           =   1575
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   54
         ToolTipText     =   "Sonido del Hechizo SFX."
         Top             =   6840
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos"
         Height          =   5895
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin VB.TextBox txtVital 
            Height          =   270
            Left            =   120
            TabIndex        =   60
            Text            =   "0"
            Top             =   1680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlStun 
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   5520
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnim 
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   4920
            Width           =   2895
         End
         Begin VB.HScrollBar scrlAnimCast 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   4320
            Width           =   2895
         End
         Begin VB.CheckBox chkAOE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hechizo de Area"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Afecta el area."
            Top             =   3240
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAOE 
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   3720
            Width           =   3015
         End
         Begin VB.HScrollBar scrlRange 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlInterval 
            Height          =   255
            Left            =   1680
            Max             =   60
            TabIndex        =   37
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDuration 
            Height          =   255
            Left            =   120
            Max             =   60
            TabIndex        =   35
            Top             =   2280
            Width           =   1455
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   1680
            Max             =   3
            TabIndex        =   22
            Top             =   480
            Width           =   1455
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1080
            Width           =   1455
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   16
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblStun 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Aturdimiento Duracion: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "El hechizo causa el efecto aturdimiento con la duracion especificada."
            Top             =   5280
            Width           =   2895
         End
         Begin VB.Label lblAnim 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Animación: Ninguna"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            ToolTipText     =   "Animacion del Hechizo. (Puedes crear mas desde el Editor de Animaciones)"
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label lblAnimCast 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Casteo Animación: Ninguna"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Animacion del Hechizo al Castearse. (Puedes crear mas desde el Editor de Animaciones)"
            Top             =   4080
            Width           =   2895
         End
         Begin VB.Label lblAOE 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Área: Asimismo"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   3480
            Width           =   3015
         End
         Begin VB.Label lblRange 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rango: Asimismo"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            ToolTipText     =   "El Area o Objetivo al cual afecta el hechizo."
            Top             =   2640
            Width           =   3015
         End
         Begin VB.Label lblInterval 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Intervalo: 0s"
            Height          =   255
            Left            =   1680
            TabIndex        =   36
            ToolTipText     =   "Cada cuantos segundos se repite el efecto del hechizo."
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblDuration 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Duracion: 0s"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            ToolTipText     =   "Duracion del Hechizo en Segundos"
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblVital 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vitalidad: "
            Height          =   255
            Left            =   120
            TabIndex        =   33
            ToolTipText     =   "Modifica el HP (Vida) del Personaje"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblDir 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dir: Arriba"
            Height          =   255
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblY 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblX 
            BackColor       =   &H00E0E0E0&
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblMap 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mapa: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informacion Basica"
         Height          =   5895
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3255
         Begin VB.PictureBox picSprite 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2640
            ScaleHeight     =   32
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   32
            TabIndex        =   49
            ToolTipText     =   "MiniImagen del Hechizo."
            Top             =   5160
            Width           =   480
         End
         Begin VB.HScrollBar scrlIcon 
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   5400
            Width           =   2415
         End
         Begin VB.HScrollBar scrlCool 
            Height          =   255
            Left            =   120
            Max             =   300
            TabIndex        =   32
            Top             =   4680
            Width           =   3015
         End
         Begin VB.HScrollBar scrlCast 
            Height          =   255
            Left            =   120
            Max             =   300
            TabIndex        =   30
            Top             =   4080
            Width           =   3015
         End
         Begin VB.ComboBox cmbClass 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   3480
            Width           =   3015
         End
         Begin VB.HScrollBar scrlAccess 
            Height          =   255
            Left            =   120
            Max             =   5
            TabIndex        =   26
            Top             =   2880
            Width           =   3015
         End
         Begin VB.HScrollBar scrlLevel 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   24
            Top             =   2280
            Width           =   3015
         End
         Begin VB.HScrollBar scrlMP 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cmbType 
            Height          =   300
            ItemData        =   "frmEditor_Spell.frx":090A
            Left            =   120
            List            =   "frmEditor_Spell.frx":0920
            Style           =   2  'Dropdown List
            TabIndex        =   11
            ToolTipText     =   "Selecciona el Tipo de Hechizo, lo cual afectara vitalmente su funcion."
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   3015
         End
         Begin VB.Label lblIcon 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Icono: Ninguno"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "MiniImagen del Hechizo."
            Top             =   5160
            Width           =   3015
         End
         Begin VB.Label lblCool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tiempo de Enfriamiento: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            ToolTipText     =   "Tiempo de Enfriamiento (Desncaso para volver a lanzar) del Hechizo."
            Top             =   4440
            Width           =   2535
         End
         Begin VB.Label lblCast 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tiempo de Casteo: Instantaneo"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            ToolTipText     =   "Tiempo que demora en lanzarse el hechizo."
            Top             =   3840
            Width           =   2415
         End
         Begin VB.Label Label5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clase Requerida: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Clase requerida para el hechizo."
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblAccess 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acceso Requerido: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Nivel de Privilegio del Jugador requerido para el lanzamiento."
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label lblLevel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel Requerido: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Nivel de Personaje requerido para lanzamiento."
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblMP 
            BackColor       =   &H00E0E0E0&
            Caption         =   "MP: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Cantidad de Puntos de Magia o MP que consume en lanzamiento."
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nombre:"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Label lblCombatLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Combate Nivel: 0"
         Height          =   180
         Left            =   4320
         TabIndex        =   59
         Top             =   6840
         Width           =   1320
      End
      Begin VB.Label lblCombatType 
         BackStyle       =   0  'Transparent
         Caption         =   "Combate Tipo:"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   6240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista de Hechizos"
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6900
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Lista de Hechizos."
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Modificar Longitud"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Agregar mas Slots."
      Top             =   7560
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()

End Sub

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If chkAOE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkarco_Click()
Spell(EditorIndex).Arco = chkarco.Value
End Sub

Private Sub chkbuff_Click()
Spell(EditorIndex).Buff = chkbuff.Value
End Sub

Private Sub chkconfusion_Click()
Spell(EditorIndex).Inversion = chkconfusion.Value
End Sub

Private Sub chkinvisibilidad_Click()
Spell(EditorIndex).Invisibilidad = chkinvisibilidad.Value
End Sub

Private Sub chkparalizar_Click()
Spell(EditorIndex).Paralisis = chkparalizar.Value
End Sub

Private Sub chksprite_Click()
Spell(EditorIndex).Sprite = chksprite.Value
End Sub

Private Sub chktransportar_Click()
Spell(EditorIndex).Transportar = chktransportar.Value
If txtx.text = "0" Then txtx.text = "1"
If txty.text = "0" Then txty.text = "1"
If txtmapa.text = "0" Then txtmapa.text = "1"
End Sub

Private Sub chkVelocidad_Click()
Spell(EditorIndex).Velocidad = chkVelocidad.Value
End Sub

Private Sub chkveneno_Click()
Spell(EditorIndex).Veneno = chkveneno.Value
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).Type = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCombatType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    Spell(EditorIndex).CombatTypeReq = cmdCombatType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCombatType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorOk False
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Label6_Click()

End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L1")
Frame1.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L2")
Frame2.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L3")
Label1.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L4")
Label2.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L5")
lblLevel.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L6")
lblAccess.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L7")
Label5.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L8")
lblCast.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L9")
lblCool.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L10")
lblIcon.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L11")
Frame3.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L12")
Frame6.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L13")
lblMap.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L14")
lblDir.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L15")
lblVital.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L16")
lblDuration.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L17")
lblInterval.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L18")
lblRange.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L19")
chkAOE.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L20")
lblAOE.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L21")
lblAnimCast.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L22")
lblAnim.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L23")
lblStun.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L24")
Frame4.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L25")
lblElement(1).Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L26")
lblElement(2).Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L27")
lblElement(3).Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L28")
chkparalizar.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L29")
chkconfusion.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L30")
chkinvisibilidad.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L31")
chkveneno.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L32")
chkVelocidad.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L33")
lblcamcorr.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L34")
lblfza.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L35")
lblInt.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L36")
lbldes.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L37")
lblVol.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L38")
lblagi.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L39")
chkbuff.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L40")
chkarco.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L41")
chktransportar.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L42")
lblMapa.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L43")
lblCombatLvl.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "L44")
lblCombatType.Caption = trad

trad = GetVar(App.Path & Lang, "SpellEditor", "B1")
cmdArray.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "B2")
cmdSave.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "B3")
cmdSSave.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "B4")
cmdDelete.Caption = trad
trad = GetVar(App.Path & Lang, "SpellEditor", "B5")
cmdCancel.Caption = trad

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L6")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = trad & " " & scrlAccess.Value
    Else
        lblAccess.Caption = trad & " Ninguno"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L22")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = trad & " " & Trim$(Animation(scrlAnim.Value).name)
    Else
        lblAnim.Caption = trad & " None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L21")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = trad & " " & Trim$(Animation(scrlAnimCast.Value).name)
    Else
        lblAnimCast.Caption = trad & " None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L20")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = trad & " " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = trad & " Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L8")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblCast.Caption = trad & " " & scrlCast.Value & "s"
    Spell(EditorIndex).CastTime = scrlCast.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCombatLvl_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L43")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_SPELLS Then Exit Sub
    lblCombatLvl.Caption = trad & " " & scrlCombatLvl
    Spell(EditorIndex).CombatLvlReq = scrlCombatLvl.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCombatLvl_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L9")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblCool.Caption = trad & ": " & scrlCool.Value & "s"
    Spell(EditorIndex).CDTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L14")

Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = trad & " " & sDir
    Spell(EditorIndex).Dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L16")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblDuration.Caption = trad & " " & scrlDuration.Value & "s"
    Spell(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlElement_Change(Index As Integer)
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L8")

Dim txt As String
    Select Case Index
        Case 1
            trad = GetVar(App.Path & Lang, "SpellEditor", "L25")
            txt = trad & " "
            Spell(EditorIndex).Dmg_Light = scrlElement(Index).Value
        Case 2
            trad = GetVar(App.Path & Lang, "SpellEditor", "L26")
            txt = trad & " "
            Spell(EditorIndex).Dmg_Dark = scrlElement(Index).Value
        Case 3
            trad = GetVar(App.Path & Lang, "SpellEditor", "L27")
            txt = trad & " "
            Spell(EditorIndex).Dmg_Neut = scrlElement(Index).Value
    End Select
    
    lblElement(Index).Caption = txt & scrlElement(Index).Value
End Sub

Private Sub scrlIcon_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L10")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = trad & " " & scrlIcon.Value
    Else
        lblIcon.Caption = trad & " Ninguno"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblInterval.Caption = "Intervalo: " & scrlInterval.Value & "s"
    Spell(EditorIndex).Interval = scrlInterval.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L5")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = trad & " " & scrlLevel.Value
    Else
        lblLevel.Caption = trad & " Ninguno"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L13")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblMap.Caption = trad & " " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "MP : " & scrlMP.Value
    Else
        lblMP.Caption = "MP : 0"
    End If
    Spell(EditorIndex).MPCost = scrlMP.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L18")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = trad & " " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = trad & " Self"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "SpellEditor", "L23")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = trad & " " & scrlStun.Value & "s"
    Else
        lblStun.Caption = trad & " 0"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblX.Caption = "X: " & scrlX.Value
    Spell(EditorIndex).X = scrlX.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblY.Caption = "Y: " & scrlY.Value
    Spell(EditorIndex).Y = scrlY.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Text5_Change()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtagi_Change()
Spell(EditorIndex).Agilidad = txtagi.text
End Sub

Private Sub txtagi_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtarco_Change()
Spell(EditorIndex).NumeroArcoItem = txtarco.text
End Sub

Private Sub txtarco_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtcaminar_Change()
Spell(EditorIndex).VelocidadCaminar2 = txtcaminar.text
End Sub

Private Sub txtcaminar_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtcorrer_Change()
Spell(EditorIndex).VelocidadCorrer2 = txtcorrer.text
End Sub

Private Sub txtcorrer_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtdes_Change()
Spell(EditorIndex).Destreza = txtdes.text
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Spell(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtfza_Change()
Spell(EditorIndex).Fuerza = txtfza.text
End Sub

Private Sub txtfza_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtint_Change()
Spell(EditorIndex).Inteligencia = txtint.text
End Sub

Private Sub txtint_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtmapa_Change()
Spell(EditorIndex).TransportarMapa = txtmapa.text
End Sub

Private Sub txtmapa_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Spell(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Spell(EditorIndex).sound = "Ninguno."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtsprite_Change()
Spell(EditorIndex).NumeroSprite = txtsprite.text
End Sub

Private Sub txtsprite_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtVeneno_Change()
Spell(EditorIndex).VenenoDmg = txtVeneno.text
End Sub

Private Sub txtVeneno_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not IsNumeric(txtVital.text) And Len(txtVital.text) > 0 Then
        txtVital.text = 1
        txtVital.SelStart = Len(txtVital.text)
    ElseIf Len(txtVital.text) < 1 Then
        Spell(EditorIndex).Vital = 1
        Exit Sub
    End If
    
    Spell(EditorIndex).Vital = CLng(txtVital.text)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtvol_Change()
Spell(EditorIndex).Voluntad = txtvol.text
End Sub

Private Sub txtvol_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtx_Change()
Spell(EditorIndex).TransportarX = txtx.text
End Sub

Private Sub txtx_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txty_Change()
Spell(EditorIndex).TransportarY = txty.text
End Sub

Private Sub txty_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then



        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub
