VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmEditor_Item 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Objetos"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12720
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   593
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   848
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame frmMaquinaObjetos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Maquina de Cubos (by ikicapitan)"
      Height          =   3375
      Left            =   3360
      TabIndex        =   141
      ToolTipText     =   $"frmEditor_Item.frx":08CA
      Top             =   4920
      Width           =   6255
      Begin VB.HScrollBar scrllDropear 
         Height          =   255
         Left            =   1320
         Max             =   500
         Min             =   1
         TabIndex        =   162
         Top             =   1440
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar scrllAnim 
         Height          =   255
         Left            =   1080
         Max             =   15
         TabIndex        =   151
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrllSFX1 
         Height          =   255
         Left            =   1080
         Max             =   15
         TabIndex        =   150
         Top             =   720
         Width           =   1815
      End
      Begin VB.HScrollBar scrllSFX2 
         Height          =   255
         Left            =   1080
         Max             =   500
         Min             =   1
         TabIndex        =   149
         Top             =   1080
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   148
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   270
         Left            =   600
         MaxLength       =   3
         TabIndex        =   147
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   270
         Left            =   320
         MaxLength       =   3
         TabIndex        =   146
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   270
         Left            =   320
         MaxLength       =   3
         TabIndex        =   145
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   144
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   143
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "reset"
         Height          =   180
         Left            =   1680
         TabIndex        =   142
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblDropear 
         BackStyle       =   0  'Transparent
         Caption         =   "Dropear: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   161
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblAnimacion 
         BackStyle       =   0  'Transparent
         Caption         =   "Anim: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   160
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblSFX1 
         BackStyle       =   0  'Transparent
         Caption         =   "SFX1: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   159
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSFX2 
         BackStyle       =   0  'Transparent
         Caption         =   "SFX2: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubo Inf Tipo: Ninguno"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   120
         TabIndex        =   156
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   155
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   154
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Golpe:"
         Height          =   255
         Left            =   1440
         TabIndex        =   153
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Dureza:"
         Height          =   255
         Left            =   1440
         TabIndex        =   152
         Top             =   4440
         Width           =   615
      End
   End
   Begin VB.Frame frmtiles 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Maquina de Cubos (EaSee Engine)"
      Height          =   5175
      Left            =   9720
      TabIndex        =   111
      ToolTipText     =   $"frmEditor_Item.frx":0957
      Top             =   3600
      Width           =   2895
      Begin VB.CommandButton reset 
         Caption         =   "reset"
         Height          =   180
         Left            =   1680
         TabIndex        =   140
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtCuboDureza 
         Height          =   270
         Left            =   2040
         MaxLength       =   3
         TabIndex        =   139
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox txtCuboGolpe 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1980
         MaxLength       =   4
         TabIndex        =   137
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox txtCuboY 
         Enabled         =   0   'False
         Height          =   270
         Left            =   320
         MaxLength       =   3
         TabIndex        =   135
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtCuboX 
         Enabled         =   0   'False
         Height          =   270
         Left            =   320
         MaxLength       =   3
         TabIndex        =   134
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox txtCuboMapa 
         Enabled         =   0   'False
         Height          =   270
         Left            =   600
         MaxLength       =   3
         TabIndex        =   133
         Top             =   4080
         Width           =   495
      End
      Begin VB.HScrollBar scrlCuboInfTipo 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   128
         Top             =   3600
         Width           =   2535
      End
      Begin VB.HScrollBar scrlCuboSupTipo 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   126
         Top             =   3000
         Width           =   2535
      End
      Begin VB.HScrollBar scrlCuboCapa2 
         Height          =   255
         Left            =   1800
         Max             =   5
         Min             =   2
         TabIndex        =   124
         Top             =   2280
         Value           =   2
         Width           =   975
      End
      Begin VB.HScrollBar scrlCuboCapa 
         Height          =   255
         Left            =   1800
         Max             =   5
         Min             =   2
         TabIndex        =   122
         Top             =   1680
         Value           =   2
         Width           =   975
      End
      Begin VB.OptionButton opt32x64 
         Caption         =   "32x64"
         Height          =   300
         Left            =   120
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   2160
         Width           =   855
      End
      Begin VB.OptionButton opt32x32 
         Caption         =   "32x32"
         Height          =   300
         Left            =   120
         MaskColor       =   &H00FFFF00&
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   1680
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.PictureBox picCubo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1200
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   118
         Top             =   1560
         Width           =   480
      End
      Begin VB.HScrollBar scrlnum 
         Height          =   255
         Left            =   1080
         Max             =   500
         Min             =   1
         TabIndex        =   117
         Top             =   1080
         Value           =   1
         Width           =   1695
      End
      Begin VB.HScrollBar scrly 
         Height          =   255
         Left            =   720
         Max             =   15
         TabIndex        =   116
         Top             =   720
         Width           =   1815
      End
      Begin VB.HScrollBar scrlx 
         Height          =   255
         Left            =   720
         Max             =   15
         TabIndex        =   115
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblCuboDureza 
         BackStyle       =   0  'Transparent
         Caption         =   "Dureza:"
         Height          =   255
         Left            =   1440
         TabIndex        =   138
         Top             =   4440
         Width           =   615
      End
      Begin VB.Label lblCuboGolpe 
         BackStyle       =   0  'Transparent
         Caption         =   "Golpe:"
         Height          =   255
         Left            =   1440
         TabIndex        =   136
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lblCuboMapaY 
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   4800
         Width           =   255
      End
      Begin VB.Label lblCuboMapaX 
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   4440
         Width           =   255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   480
         TabIndex        =   130
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblCuboMapa 
         BackStyle       =   0  'Transparent
         Caption         =   "Mapa:"
         Height          =   255
         Left            =   120
         TabIndex        =   129
         Top             =   4080
         Width           =   495
      End
      Begin VB.Label lblCuboInfTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubo Inf Tipo: Ninguno"
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblCuboSupTipo 
         BackStyle       =   0  'Transparent
         Caption         =   "Cubo Sup Tipo: Ninguno"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label lblCuboCapa2 
         BackStyle       =   0  'Transparent
         Caption         =   "Capa 2: 1"
         Height          =   255
         Left            =   1800
         TabIndex        =   123
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblCuboCapa 
         BackStyle       =   0  'Transparent
         Caption         =   "Capa 1: 1"
         Height          =   255
         Left            =   1800
         TabIndex        =   121
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblnumero 
         BackStyle       =   0  'Transparent
         Caption         =   "Numero: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   114
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lbly 
         BackStyle       =   0  'Transparent
         Caption         =   "Y: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   113
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblx 
         BackStyle       =   0  'Transparent
         Caption         =   "X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proyectiles"
      Height          =   3375
      Left            =   9720
      TabIndex        =   101
      ToolTipText     =   "En caso de ser un arco, pistola o arma de lanzamiento de proyectiles aqui puedes configurar la imagen del proyectil y demas."
      Top             =   120
      Width           =   2895
      Begin VB.HScrollBar scrlmunicion 
         Height          =   255
         Left            =   120
         TabIndex        =   165
         Top             =   3000
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectileRange 
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   2400
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectileSpeed 
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectileDamage 
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   1200
         Width           =   2655
      End
      Begin VB.HScrollBar scrlProjectilePic 
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Lblmunicion 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Municion: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   164
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label lblProjectileRange 
         BackStyle       =   0  'Transparent
         Caption         =   "Rango: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label lblProjectileSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Velocidad: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   105
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label lblProjectileDamage 
         BackStyle       =   0  'Transparent
         Caption         =   "Golpe: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblProjectilePic 
         BackStyle       =   0  'Transparent
         Caption         =   "Img: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   103
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame fraBooks 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Libro"
      Height          =   3375
      Left            =   3360
      TabIndex        =   99
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin RichTextLib.RichTextBox rtbBookText2 
         Height          =   3015
         Left            =   3240
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         MaxLength       =   1000
         TextRTF         =   $"frmEditor_Item.frx":09FF
      End
      Begin RichTextLib.RichTextBox rtbBookText 
         Height          =   3015
         Left            =   600
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5318
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         MaxLength       =   1000
         TextRTF         =   $"frmEditor_Item.frx":0A79
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisitos"
      Height          =   1335
      Left            =   3360
      TabIndex        =   6
      ToolTipText     =   "Requisitos para poder utilizar este Objeto."
      Top             =   3600
      Width           =   6255
      Begin VB.ComboBox cmbSkill 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   4080
         Max             =   255
         TabIndex        =   83
         Top             =   960
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   5160
         Max             =   255
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2880
         Max             =   255
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.Label label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   180
         Left            =   120
         TabIndex        =   84
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   390
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Nivel: 0"
         Height          =   180
         Index           =   6
         Left            =   2880
         TabIndex        =   82
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fza: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Res: 0"
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vol: 0"
         Height          =   180
         Index           =   5
         Left            =   2280
         TabIndex        =   12
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   6120
      TabIndex        =   86
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Propiedades de Objeto"
      Height          =   3375
      Left            =   3360
      TabIndex        =   17
      Top             =   120
      Width           =   6255
      Begin VB.TextBox txtDesc 
         Height          =   735
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         ToolTipText     =   "Descripcion legible del objeto."
         Top             =   1800
         Width           =   2655
      End
      Begin VB.HScrollBar scrlCombatLvl 
         Height          =   255
         LargeChange     =   10
         Left            =   1680
         Max             =   100
         TabIndex        =   77
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox cmdCombatType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0AF3
         Left            =   1320
         List            =   "frmEditor_Item.frx":0B12
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CheckBox chkStackable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Apilable"
         Height          =   255
         Left            =   2880
         TabIndex        =   75
         Top             =   3000
         Width           =   1335
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   4200
         Max             =   99
         TabIndex        =   73
         Top             =   2760
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   71
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0B8B
         Left            =   3840
         List            =   "frmEditor_Item.frx":0B8D
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cmbSound 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0B8F
         Left            =   3720
         List            =   "frmEditor_Item.frx":0B91
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   1680
         Width           =   2415
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   4200
         Max             =   5
         TabIndex        =   25
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cmbBind 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0B93
         Left            =   4200
         List            =   "frmEditor_Item.frx":0BA0
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   4200
         Max             =   30000
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   5040
         Max             =   5
         TabIndex        =   22
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0BD1
         Left            =   120
         List            =   "frmEditor_Item.frx":0C0B
         Style           =   2  'Dropdown List
         TabIndex        =   21
         ToolTipText     =   "Clase o Tipo de Objeto. Esto afecta su funcion al ser utilizado."
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   960
         Max             =   1800
         TabIndex        =   19
         Top             =   600
         Value           =   1600
         Width           =   1215
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   18
         ToolTipText     =   "Imagen del Objeto."
         Top             =   600
         Width           =   480
      End
      Begin VB.Label lblCombatLvl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   3000
         Width           =   1245
      End
      Begin VB.Label lblCombatType 
         BackStyle       =   0  'Transparent
         Caption         =   "Combat Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   74
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso Req: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   72
         Top             =   2400
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clase Req:"
         Height          =   180
         Left            =   2880
         TabIndex        =   70
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sonido:"
         Height          =   255
         Left            =   2880
         TabIndex        =   67
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rareza: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   31
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unido Tipo:"
         Height          =   180
         Left            =   2880
         TabIndex        =   30
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio: 0"
         Height          =   180
         Left            =   2880
         TabIndex        =   29
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Animacion: None"
         Height          =   180
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Img: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Modificar Longitud"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   8400
      Width           =   2895
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   8400
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   8400
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   7320
      TabIndex        =   2
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista Objetos"
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7800
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraSpell 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Spell Data"
      Height          =   1215
      Left            =   3360
      TabIndex        =   52
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
      Begin VB.HScrollBar scrlSpell 
         Height          =   255
         Left            =   1080
         Max             =   255
         Min             =   1
         TabIndex        =   53
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblSpellName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: None"
         Height          =   180
         Left            =   240
         TabIndex        =   55
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblSpell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Frame fraEquipment 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equipment Data"
      Height          =   3375
      Left            =   3360
      TabIndex        =   32
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   11
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   97
         Top             =   2640
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   10
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   95
         Top             =   2640
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   8
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   93
         Top             =   2280
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   7
         LargeChange     =   10
         Left            =   4440
         Max             =   255
         TabIndex        =   91
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   6
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   89
         Top             =   1920
         Width           =   1095
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   9
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   87
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkHanded 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Two-Handed"
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   3000
         Width           =   1335
      End
      Begin VB.PictureBox picPaperdoll 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5640
         ScaleHeight     =   56
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   58
         Top             =   2040
         Width           =   480
      End
      Begin VB.HScrollBar scrlPaperdoll 
         Height          =   255
         Left            =   4200
         TabIndex        =   57
         Top             =   3000
         Width           =   1335
      End
      Begin VB.HScrollBar scrlSpeed 
         Height          =   255
         LargeChange     =   100
         Left            =   4560
         Max             =   3000
         Min             =   100
         SmallChange     =   100
         TabIndex        =   40
         Top             =   840
         Value           =   100
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   39
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   38
         Top             =   1560
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   4680
         Max             =   255
         TabIndex        =   37
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   3000
         Max             =   255
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.HScrollBar scrlDamage 
         Height          =   255
         LargeChange     =   10
         Left            =   1320
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbTool 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":0CA6
         Left            =   1320
         List            =   "frmEditor_Item.frx":0CB6
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   4815
      End
      Begin VB.HScrollBar scrlStatBonus 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   960
         Max             =   255
         TabIndex        =   33
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Resist: 0"
         Height          =   180
         Index           =   11
         Left            =   2880
         TabIndex        =   98
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Neutral Dmg: 0"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   96
         Top             =   2640
         UseMnemonic     =   0   'False
         Width           =   1155
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Resist: 0"
         Height          =   180
         Index           =   8
         Left            =   2880
         TabIndex        =   94
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Resist: 0"
         Height          =   180
         Index           =   7
         Left            =   2880
         TabIndex        =   92
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Light Dmg: 0"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   90
         Top             =   1920
         UseMnemonic     =   0   'False
         Width           =   975
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dark Dmg: 0"
         Height          =   180
         Index           =   9
         Left            =   120
         TabIndex        =   88
         Top             =   2280
         UseMnemonic     =   0   'False
         Width           =   960
      End
      Begin VB.Label lblPaperdoll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paperdoll: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   56
         Top             =   3000
         Width           =   915
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed: 0.1 sec"
         Height          =   180
         Left            =   3240
         TabIndex        =   48
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2160
         TabIndex        =   47
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   630
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   615
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Int: 0"
         Height          =   180
         Index           =   3
         Left            =   3960
         TabIndex        =   45
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ End: 0"
         Height          =   180
         Index           =   2
         Left            =   2160
         TabIndex        =   44
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   645
      End
      Begin VB.Label lblDamage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Damage: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Object Tool:"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblStatBonus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+ Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   41
         Top             =   1200
         UseMnemonic     =   0   'False
         Width           =   585
      End
   End
   Begin VB.Frame fraVitals 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consume Data"
      Height          =   3135
      Left            =   3360
      TabIndex        =   49
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CheckBox chkInstant 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Instant Cast?"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.HScrollBar scrlCastSpell 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   64
         Top             =   2400
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddExp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   62
         Top             =   1800
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddMP 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   1200
         Width           =   3495
      End
      Begin VB.HScrollBar scrlAddHp 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   50
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblCastSpell 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cast Spell: None"
         Height          =   180
         Left            =   120
         TabIndex        =   65
         Top             =   2160
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblAddExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add Exp: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   63
         Top             =   1560
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label lblAddMP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add MP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   960
         UseMnemonic     =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddHP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add HP: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   51
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   780
      End
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easee Engine"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4695
      Left            =   9840
      TabIndex        =   102
      Top             =   3840
      Width           =   2655
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long





Private Sub Check1_Click()

End Sub





Private Sub chkHanded_Click()
    Item(EditorIndex).Handed = chkHanded.Value
End Sub

Private Sub cmbBind_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).BindType = cmbBind.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbBind_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSkill_Click()
    Item(EditorIndex).SkillReq = cmbSkill.ListIndex + 1
End Sub

Private Sub cmbSound_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If cmbSound.ListIndex >= 0 Then
        Item(EditorIndex).sound = cmbSound.List(cmbSound.ListIndex)
    Else
        Item(EditorIndex).sound = "Ninguno."
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSound_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCombatType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).CombatTypeReq = cmdCombatType.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCombatType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorOk(False)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub





Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "C1")
Me.Caption = trad

trad = GetVar(App.Path & Lang, "ItemEditor", "L1")
Frame3.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L2")
Frame2.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L3")
Label1.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L4")
Label3.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L5")
lblPrice.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L6")
Label11.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L7")
lblRarity.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L8")
lblAnim.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L9")
Label4.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L10")
Label2.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L11")
lblAccessReq.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L12")
lblCombatType.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L13")
lblCombatLvl.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L14")
lblLevelReq.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L15")
chkStackable.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L16")
Frame1.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L17")
lblStatReq(1).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L18")
lblStatReq(2).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L19")
lblStatReq(3).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L20")
lblStatReq(4).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L21")
lblStatReq(5).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L22")
label.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L23")
lblStatReq(6).Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L24")
frmMaquinaObjetos.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L25")
lblAnimacion.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L26")
lblSFX1.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L27")
lblSFX2.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L28")
lblDropear.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L29")
Frame4.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L30")
lblProjectileDamage.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L31")
lblProjectileSpeed.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L32")
lblProjectileRange.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L33")
Lblmunicion.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L34")
frmtiles.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L35")
lblnumero.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L36")
lblCuboCapa.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L37")
lblCuboCapa2.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L38")
lblCuboSupTipo.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L39")
lblCuboInfTipo.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L40")
lblCuboMapa.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L41")
lblCuboGolpe.Caption = trad
trad = GetVar(App.Path & Lang, "ItemEditor", "L42")
lblCuboDureza.Caption = trad


Dim I As Integer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    scrlPic.max = numitems
    scrlAnim.max = MAX_ANIMATIONS
    scrlPaperdoll.max = NumPaperdolls
    scrlDamage.max = MAX_INTEGER
    scrlmunicion.max = MAX_ITEMS

    
    'set main txt for books
    rtbBookText.text = "Input book text here." & vbNewLine & _
    "Use /t to automatically tab your text 4 spaces."
    rtbBookText.MaxLength = TEXT_LENGTH
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    If (cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (cmbType.ListIndex = ITEM_TYPE_PICACUBOS) Then
        fraEquipment.Visible = True
        chkStackable.Visible = False
        Item(EditorIndex).Stackable = 0
        'scrlDamage_Change
        If cmbType.ListIndex = ITEM_TYPE_WEAPON Or cmbType.ListIndex = ITEM_TYPE_PICACUBOS Then
            Frame4.Visible = True
            Me.Width = 12825
             
        End If
     
    Else
        fraEquipment.Visible = False
        Frame4.Visible = False
        Me.Width = 9810
        
        chkStackable.Visible = True
    End If

    If cmbType.ListIndex = ITEM_TYPE_CONSUME Then
        fraVitals.Visible = True
        'scrlVitalMod_Change
    Else
        fraVitals.Visible = False
    End If

    If (cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        fraSpell.Visible = True
    Else
        fraSpell.Visible = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_WEAPON Or cmbType.ListIndex = ITEM_TYPE_SPELL Or cmbType.ListIndex = ITEM_TYPE_PICACUBOS Then
        cmdCombatType.Enabled = True
        scrlCombatLvl.Enabled = True
    Else
        cmdCombatType.Enabled = False
        scrlCombatLvl.Enabled = False
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_BOOK Then
        fraBooks.Visible = True
        rtbBookText.text = "Erase una vez EaSee Engine"
    Else
        fraBooks.Visible = False
        rtbBookText.text = "El mejor Engine de todos los tiempos"
    End If
    
    If cmbType.ListIndex = ITEM_TYPE_CUBO Then 'Item Proyecto Comunitario EaSee Bloque o Cubo
       Me.Width = 12825
        frmtiles.Visible = True
        frmMaquinaObjetos.Visible = True
        Else
         frmtiles.Visible = False
         frmMaquinaObjetos.Visible = False

    End If
    
    Item(EditorIndex).Type = cmbType.ListIndex
   
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkStackable_Click()
    Item(EditorIndex).Stackable = chkStackable.Value
End Sub





Private Sub HScroll7_Change()

End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Option1_Click()

End Sub

Private Sub opt32x32_Click()
opt32x32.Value = True
opt32x64.Value = False
picCubo.Height = 480
Item(EditorIndex).Cubo32 = True
Item(EditorIndex).Cubo64 = False
End Sub

Private Sub opt32x64_Click()
opt32x32.Value = False
opt32x64.Value = True
picCubo.Height = 960
Item(EditorIndex).Cubo32 = False
Item(EditorIndex).Cubo64 = True
End Sub

Private Sub reset_Click()
Dim I As Integer
For I = 1 To MAX_ITEMS
Item(I).CuboCapa1 = 2
Item(I).CuboCapa2 = 1
Item(I).CuboInfTipo = 0
Item(I).CuboSupTipo = 0
Item(I).CuboTileN = 1
Next
End Sub

Private Sub rtbBookText_Change()
Dim tLoc As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not frmEditor_Item.Visible = True Then Exit Sub
    
    tLoc = InStr(rtbBookText.text, "/t")
    If tLoc > 0 Then
        rtbBookText.text = Replace$(rtbBookText.text, "/t", "    ")
        rtbBookText.SelStart = tLoc + 3
    End If
    
    If Len(rtbBookText.text) > 0 Then
        Item(EditorIndex).Book.name = Item(EditorIndex).name
        Item(EditorIndex).Book.text = rtbBookText.text
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "rtbBookText_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
Private Sub rtbBookText2_Change()
Dim tLoc As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not frmEditor_Item.Visible = True Then Exit Sub
    
    tLoc = InStr(rtbBookText2.text, "/t")
    If tLoc > 0 Then
        rtbBookText2.text = Replace$(rtbBookText2.text, "/t", "    ")
        rtbBookText2.SelStart = tLoc + 3
    End If
    
    If Len(rtbBookText2.text) > 0 Then
        Item(EditorIndex).Book.Text2 = rtbBookText2.text
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "rtbBookText2_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub scrlAccessReq_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L11")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = trad & " " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddHP.Caption = "Agregar HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddMP.Caption = "Agregar MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblAddExp.Caption = "Agregar Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L8")

Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).name)
    End If
    lblAnim.Caption = trad & " " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCombatLvl_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L13")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblCombatLvl.Caption = trad & " " & scrlCombatLvl
    Item(EditorIndex).CombatLvlReq = scrlCombatLvl.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlCombatLvl_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCuboCapa_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L36")
lblCuboCapa.Caption = trad & " " & scrlCuboCapa.Value
Item(EditorIndex).CuboCapa1 = scrlCuboCapa.Value
End Sub

Private Sub scrlCuboCapa2_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L37")
lblCuboCapa2.Caption = trad & " " & scrlCuboCapa2.Value
Item(EditorIndex).CuboCapa2 = scrlCuboCapa2.Value
End Sub

Private Sub scrlCuboInfTipo_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L39")
Dim Valor2 As Long
Valor2 = scrlCuboInfTipo.Value

Select Case Valor2

Case 0:
lblCuboInfTipo.Caption = trad & " None"

Case 1:
lblCuboInfTipo.Caption = trad & " Block"

Case 2:
lblCuboInfTipo.Caption = trad & " Chest"

Case 3:
lblCuboInfTipo.Caption = trad & " Warp"

Case 4:
lblCuboInfTipo.Caption = trad & " Trap"

Case 5:
lblCuboInfTipo.Caption = trad & " Message"

End Select
Item(EditorIndex).CuboInfTipo = scrlCuboInfTipo.Value

If Valor2 = 3 Then
txtCuboMapa.Enabled = True
txtCuboMapa.BackColor = &H80000005
txtCuboX.Enabled = True
txtCuboX.BackColor = &H80000005
txtCuboY.Enabled = True
txtCuboY.BackColor = &H80000005
Else
If scrlCuboSupTipo.Value = 3 Then

Else
txtCuboMapa.Enabled = False
txtCuboMapa.BackColor = &H0&
txtCuboX.Enabled = False
txtCuboX.BackColor = &H0&
txtCuboY.Enabled = False
txtCuboY.BackColor = &H0&
End If
End If

If Valor2 = 4 Then
txtCuboGolpe.Enabled = True
txtCuboGolpe.BackColor = &H80000005
Else
If scrlCuboSupTipo.Value = 4 Then

Else
txtCuboGolpe.Enabled = False
txtCuboGolpe.BackColor = &H0&
End If
End If

End Sub

Private Sub scrlCuboSupTipo_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L38")
Dim Valor As Long
Valor = scrlCuboSupTipo.Value

Select Case Valor

Case 0:
lblCuboSupTipo.Caption = trad & " None"

Case 1:
lblCuboSupTipo.Caption = trad & " Block"

Case 2:
lblCuboSupTipo.Caption = trad & " Chest"

Case 3:
lblCuboSupTipo.Caption = trad & " Warp"

Case 4:
lblCuboSupTipo.Caption = trad & " Trap"

Case 5:
lblCuboSupTipo.Caption = trad & " Message"

End Select
Item(EditorIndex).CuboSupTipo = Valor

If Valor = 3 Then
txtCuboMapa.Enabled = True
txtCuboMapa.BackColor = &H80000005
txtCuboX.Enabled = True
txtCuboX.BackColor = &H80000005
txtCuboY.Enabled = True
txtCuboY.BackColor = &H80000005
Else
If scrlCuboInfTipo.Value = 3 Then

Else
txtCuboMapa.Enabled = False
txtCuboMapa.BackColor = &H0&
txtCuboX.Enabled = False
txtCuboX.BackColor = &H0&
txtCuboY.Enabled = False
txtCuboY.BackColor = &H0&
End If
End If

If Valor = 4 Then
txtCuboGolpe.Enabled = True
txtCuboGolpe.BackColor = &H80000005
Else
If scrlCuboInfTipo.Value = 4 Then

Else
txtCuboGolpe.Enabled = False
txtCuboGolpe.BackColor = &H0&
End If
End If

End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Golpe: " & scrlDamage.Value
    Item(EditorIndex).Data2 = scrlDamage.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub


Private Sub scrllAnim_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L25")

lblAnimacion.Caption = trad & " " & scrllAnim.Value
Item(EditorIndex).CuboAnimacion = scrllAnim.Value

End Sub

Private Sub scrlLevelReq_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L14")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = trad & " " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub



Private Sub scrllSFX1_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L26")

lblSFX1.Caption = trad & " " & scrllSFX1.Value
Item(EditorIndex).CuboSFX1 = scrllSFX1.Value

End Sub

Private Sub scrllSFX2_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L27")
lblSFX2.Caption = trad & " " & scrllSFX2.Value
Item(EditorIndex).CuboSFX2 = scrllSFX2.Value

End Sub

Private Sub scrlmunicion_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L33")
Lblmunicion.Caption = trad & " " & scrlmunicion.Value
Item(EditorIndex).ProjecTile.Municion = scrlmunicion.Value
End Sub

Private Sub scrlNum_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L35")
If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
lblnumero.Caption = trad & " " & scrlNum.Value
Item(EditorIndex).CuboTileN = scrlNum.Value

    Exit Sub
ErrorHandler:
    HandleError "scrlNum_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub

End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Img: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L5")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = trad & " " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L7")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = trad & " " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Velocidad: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
        Case 6
            text = "Light Dmg: "
            Item(EditorIndex).Element_Light_Dmg = scrlStatBonus(Index).Value
        Case 7
            text = "Light Resist: "
            Item(EditorIndex).Element_Light_Res = scrlStatBonus(Index).Value
        Case 8
            text = "Dark Resist: "
            Item(EditorIndex).Element_Dark_Res = scrlStatBonus(Index).Value
        Case 9
            text = "Dark Dmg: "
            Item(EditorIndex).Element_Dark_Dmg = scrlStatBonus(Index).Value
        Case 10
            text = "Neutral Dmg: "
            Item(EditorIndex).Element_Neut_Dmg = scrlStatBonus(Index).Value
        Case 11
            text = "Neutral Resist: "
            Item(EditorIndex).Element_Neut_Res = scrlStatBonus(Index).Value
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    If Index < 6 Then Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L5")
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
        trad = GetVar(App.Path & Lang, "ItemEditor", "L17")
            text = trad & " "
        Case 2
        trad = GetVar(App.Path & Lang, "ItemEditor", "L18")
            text = trad & " "
        Case 3
        trad = GetVar(App.Path & Lang, "ItemEditor", "L19")
            text = trad & " "
        Case 4
        trad = GetVar(App.Path & Lang, "ItemEditor", "L20")
            text = trad & " "
        Case 5
        trad = GetVar(App.Path & Lang, "ItemEditor", "L21")
            text = trad & " "
        Case 6
        trad = GetVar(App.Path & Lang, "ItemEditor", "L23")
            text = trad & " "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If Len(Trim$(Spell(scrlSpell.Value).name)) > 0 Then
        lblSpellName.Caption = "Nombre: " & Trim$(Spell(scrlSpell.Value).name)
    Else
        lblSpellName.Caption = "Nombre: Ninguno"
    End If
    
    lblSpell.Caption = "Hechizo: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()

lblX.Caption = "X: " & scrlX.Value
Item(EditorIndex).CuboTileX = scrlX.Value

End Sub

Private Sub scrlY_Change()
lblY.Caption = "Y: " & scrlY.Value
Item(EditorIndex).CuboTileY = scrlY.Value
End Sub

Private Sub scrllDropear_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L28")
lblDropear.Caption = trad & " " & scrllDropear.Value
Item(EditorIndex).CuboObjeto = scrllDropear.Value

End Sub

Private Sub txtCuboDureza_Change()
Item(EditorIndex).CuboDureza = CInt(txtCuboDureza.text)
End Sub

Private Sub txtCuboGolpe_Change()
Item(EditorIndex).CuboGolpe = CInt(txtCuboGolpe.text)
End Sub

Private Sub txtCuboMapa_Change()
If CInt(txtCuboMapa.text) > MAX_MAPS Then
txtCuboMapa.text = MAX_MAPS 'Para que no exceda el maximo de mapas y genere error
End If
Item(EditorIndex).CuboMapa = CInt(txtCuboMapa.text)
End Sub

Private Sub txtCuboX_Change()
Item(EditorIndex).CuboMapaX = CInt(txtCuboX.text)
End Sub

Private Sub txtCuboY_Change()
Item(EditorIndex).CuboMapaY = CInt(txtCuboY.text)
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileDamage_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L30")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileDamage.Caption = trad & " " & scrlProjectileDamage.Value
    Item(EditorIndex).ProjecTile.Damage = scrlProjectileDamage.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectilePic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectilePic.Caption = "Img: " & scrlProjectilePic.Value
    Item(EditorIndex).ProjecTile.Pic = scrlProjectilePic.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectilePic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ProjecTile
Private Sub scrlProjectileRange_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L32")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileRange.Caption = trad & " " & scrlProjectileRange.Value
    Item(EditorIndex).ProjecTile.Range = scrlProjectileRange.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlProjectileRange_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' projectile
Private Sub scrlProjectileSpeed_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "ItemEditor", "L31")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProjectileSpeed.Caption = trad & " " & scrlProjectileSpeed.Value
    Item(EditorIndex).ProjecTile.speed = scrlProjectileSpeed.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
