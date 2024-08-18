VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmEditor_Events 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Eventos"
   ClientHeight    =   8970
   ClientLeft      =   855
   ClientTop       =   1755
   ClientWidth     =   12855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditor_Events.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   857
   Begin VB.Frame fraGraphic 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccion de Imagen"
      Height          =   375
      Left            =   120
      TabIndex        =   74
      Top             =   120
      Visible         =   0   'False
      Width           =   495
      Begin VB.HScrollBar hScrlGraphicSel 
         Height          =   255
         LargeChange     =   64
         Left            =   240
         SmallChange     =   32
         TabIndex        =   105
         Top             =   7920
         Visible         =   0   'False
         Width           =   11895
      End
      Begin VB.VScrollBar vScrlGraphicSel 
         Height          =   7095
         LargeChange     =   64
         Left            =   12240
         SmallChange     =   32
         TabIndex        =   104
         Top             =   720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picGraphicSel 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7080
         Left            =   240
         ScaleHeight     =   472
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   793
         TabIndex        =   81
         Top             =   720
         Width           =   11895
      End
      Begin VB.CommandButton cmdGraphicCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11040
         TabIndex        =   80
         Top             =   8280
         Width           =   1455
      End
      Begin VB.CommandButton cmdGraphicOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   79
         Top             =   8280
         Width           =   1455
      End
      Begin VB.ComboBox cmbGraphic 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":08CA
         Left            =   720
         List            =   "frmEditor_Events.frx":08D7
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   240
         Width           =   2175
      End
      Begin VB.HScrollBar scrlGraphic 
         Height          =   255
         Left            =   4440
         TabIndex        =   75
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   33
         Left            =   120
         TabIndex        =   78
         Top             =   270
         Width           =   855
      End
      Begin VB.Label lblGraphic 
         Caption         =   "Numero: 1"
         Height          =   255
         Left            =   3000
         TabIndex        =   77
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraLabeling 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetando Variables y Switches"
      Height          =   495
      Left            =   120
      TabIndex        =   321
      Top             =   120
      Visible         =   0   'False
      Width           =   615
      Begin VB.Frame fraRenaming 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Renombrando Variable/Switch"
         Height          =   8535
         Left            =   120
         TabIndex        =   330
         Top             =   120
         Visible         =   0   'False
         Width           =   12615
         Begin VB.Frame fraRandom 
            Caption         =   "Editar Variable/Switch"
            Height          =   2295
            Index           =   10
            Left            =   3600
            TabIndex        =   331
            Top             =   2520
            Width           =   5055
            Begin VB.TextBox txtRename 
               Height          =   375
               Left            =   120
               TabIndex        =   334
               Top             =   720
               Width           =   4815
            End
            Begin VB.CommandButton cmdRename_Cancel 
               Caption         =   "Cancelar"
               Height          =   375
               Left            =   3720
               TabIndex        =   333
               Top             =   1800
               Width           =   1215
            End
            Begin VB.CommandButton cmdRename_Ok 
               Caption         =   "Ok"
               Height          =   375
               Left            =   2280
               TabIndex        =   332
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label lblEditing 
               Caption         =   "Nombrando Variable #1"
               Height          =   375
               Left            =   120
               TabIndex        =   335
               Top             =   360
               Width           =   4815
            End
         End
      End
      Begin VB.CommandButton cmdRenameSwitch 
         Caption         =   "Renombrar Switch"
         Height          =   375
         Left            =   8280
         TabIndex        =   329
         Top             =   7320
         Width           =   1935
      End
      Begin VB.CommandButton cmdRenameVariable 
         Caption         =   "Renombrar Variable"
         Height          =   375
         Left            =   360
         TabIndex        =   328
         Top             =   7320
         Width           =   1935
      End
      Begin VB.ListBox lstSwitches 
         Height          =   6495
         Left            =   8280
         TabIndex        =   326
         Top             =   720
         Width           =   3855
      End
      Begin VB.ListBox lstVariables 
         Height          =   6495
         Left            =   360
         TabIndex        =   324
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmbLabel_Ok 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   323
         Top             =   8400
         Width           =   1455
      End
      Begin VB.CommandButton cmdLabel_Cancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11040
         TabIndex        =   322
         Top             =   8400
         Width           =   1455
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Switches Jugador"
         Height          =   255
         Index           =   36
         Left            =   8280
         TabIndex        =   327
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblRandomLabel 
         Caption         =   "Variables Jugador"
         Height          =   255
         Index           =   25
         Left            =   360
         TabIndex        =   325
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraMoveRoute 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ruta de Movimiento"
      Height          =   375
      Left            =   120
      TabIndex        =   106
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      Begin VB.Frame fraRandom 
         Caption         =   "Comandos"
         Height          =   6615
         Index           =   14
         Left            =   3120
         TabIndex        =   113
         Top             =   480
         Width           =   9255
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Sprite..."
            Height          =   375
            Index           =   42
            Left            =   6720
            TabIndex        =   156
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Posic Encima del Jugador"
            Height          =   375
            Index           =   41
            Left            =   6720
            TabIndex        =   155
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Posicion con Jugador"
            Height          =   375
            Index           =   40
            Left            =   6720
            TabIndex        =   154
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Posicion debajo de jugador"
            Height          =   375
            Index           =   39
            Left            =   6720
            TabIndex        =   153
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "No atravieza"
            Height          =   375
            Index           =   38
            Left            =   6720
            TabIndex        =   152
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Pasar a Traves"
            Height          =   375
            Index           =   37
            Left            =   6720
            TabIndex        =   151
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "No Corregir Dir"
            Height          =   375
            Index           =   36
            Left            =   6720
            TabIndex        =   150
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Corregir Dir"
            Height          =   375
            Index           =   35
            Left            =   4560
            TabIndex        =   149
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Sin Animacion"
            Height          =   375
            Index           =   34
            Left            =   4560
            TabIndex        =   148
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Animacion Caminar"
            Height          =   375
            Index           =   33
            Left            =   4560
            TabIndex        =   147
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Frecuencia Mayor"
            Height          =   375
            Index           =   32
            Left            =   4560
            TabIndex        =   146
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Frecuencia Alta"
            Height          =   375
            Index           =   31
            Left            =   4560
            TabIndex        =   145
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Frecuencia Normal"
            Height          =   375
            Index           =   30
            Left            =   4560
            TabIndex        =   144
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Frecuencia Menor"
            Height          =   375
            Index           =   29
            Left            =   4560
            TabIndex        =   143
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Frecuencia Baja"
            Height          =   375
            Index           =   28
            Left            =   4560
            TabIndex        =   142
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Acelerar x4"
            Height          =   375
            Index           =   27
            Left            =   4560
            TabIndex        =   141
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Acelerar x2"
            Height          =   375
            Index           =   26
            Left            =   4560
            TabIndex        =   140
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Velocidad Normal"
            Height          =   375
            Index           =   25
            Left            =   4560
            TabIndex        =   139
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Ralentizar x2"
            Height          =   375
            Index           =   24
            Left            =   4560
            TabIndex        =   138
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Ralentizar x4"
            Height          =   375
            Index           =   23
            Left            =   2400
            TabIndex        =   137
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Ralentizar x8"
            Height          =   375
            Index           =   22
            Left            =   2400
            TabIndex        =   136
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Dar Espalda a Jugador"
            Height          =   375
            Index           =   21
            Left            =   2400
            TabIndex        =   135
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Girar Hacia Jugador"
            Height          =   375
            Index           =   20
            Left            =   2400
            TabIndex        =   134
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Girar al Azar"
            Height          =   375
            Index           =   19
            Left            =   2400
            TabIndex        =   133
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Girar 180°"
            Height          =   375
            Index           =   18
            Left            =   2400
            TabIndex        =   132
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Girar 90° a la Izquierda"
            Height          =   375
            Index           =   17
            Left            =   2400
            TabIndex        =   131
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Girar 90° a la Derecha"
            Height          =   375
            Index           =   16
            Left            =   2400
            TabIndex        =   130
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mirar Der"
            Height          =   375
            Index           =   15
            Left            =   2400
            TabIndex        =   129
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mirar Izq"
            Height          =   375
            Index           =   14
            Left            =   2400
            TabIndex        =   128
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mirar Abajo"
            Height          =   375
            Index           =   13
            Left            =   2400
            TabIndex        =   127
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mirar Arriba"
            Height          =   375
            Index           =   12
            Left            =   2400
            TabIndex        =   126
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Esperar 1s"
            Height          =   375
            Index           =   11
            Left            =   240
            TabIndex        =   125
            Top             =   5640
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Esperar 1/2s"
            Height          =   375
            Index           =   10
            Left            =   240
            TabIndex        =   124
            Top             =   5160
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Esperar 100Ms"
            Height          =   375
            Index           =   9
            Left            =   240
            TabIndex        =   123
            Top             =   4680
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Paso Atras"
            Height          =   375
            Index           =   8
            Left            =   240
            TabIndex        =   122
            Top             =   4200
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Paso Adelante"
            Height          =   375
            Index           =   7
            Left            =   240
            TabIndex        =   121
            Top             =   3720
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Alejarse del Jugador"
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   120
            Top             =   3240
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Ir al Jugador"
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   119
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mover azar"
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   118
            Top             =   2280
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mover Derecha"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   117
            Top             =   1800
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mover Izq"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   116
            Top             =   1320
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mover Abajo"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   115
            Top             =   840
            Width           =   1935
         End
         Begin VB.CommandButton cmdAddMoveRoute 
            Caption         =   "Mover Arriba"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   114
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblRandomLabel 
            Caption         =   "*** No se procesan en eventos globales."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   157
            Top             =   6240
            Width           =   8535
         End
      End
      Begin VB.ComboBox cmbEvent 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":08F8
         Left            =   120
         List            =   "frmEditor_Events.frx":08FA
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chkRepeatRoute 
         Caption         =   "Repetir Ruta"
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   7560
         Width           =   2655
      End
      Begin VB.CheckBox chkIgnoreMove 
         Caption         =   "Ignorar bloqueos."
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   7200
         Width           =   2655
      End
      Begin VB.ListBox lstMoveRoute 
         Height          =   6105
         Left            =   120
         TabIndex        =   109
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton cmdMoveRouteOk 
         Caption         =   "OK"
         Height          =   375
         Left            =   9480
         TabIndex        =   108
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CommandButton cmdMoveRouteCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11040
         TabIndex        =   107
         Top             =   8160
         Width           =   1455
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Posicionamiento"
      Height          =   855
      Index           =   19
      Left            =   2760
      TabIndex        =   102
      ToolTipText     =   $"frmEditor_Events.frx":08FC
      Top             =   5880
      Width           =   3375
      Begin VB.ComboBox cmbPositioning 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":09AF
         Left            =   120
         List            =   "frmEditor_Events.frx":09BC
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Global"
      Height          =   615
      Index           =   17
      Left            =   2760
      TabIndex        =   99
      Top             =   7680
      Width           =   3375
      Begin VB.CheckBox chkGlobal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Global**"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   9720
      TabIndex        =   36
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11280
      TabIndex        =   35
      Top             =   8520
      Width           =   1455
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General"
      Height          =   735
      Index           =   20
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   12615
      Begin VB.CommandButton cmdClearPage 
         Caption         =   "Limpiar Página"
         Height          =   375
         Left            =   10920
         TabIndex        =   33
         ToolTipText     =   "Elimina todos los cambios hechos en la pagina actual del Evento actual."
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeletePage 
         Caption         =   "Eliminar Página"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9360
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPastePage 
         Caption         =   "Pegar Página"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7800
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCopyPage 
         Caption         =   "Copiar Página"
         Height          =   375
         Left            =   6240
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdNewPage 
         Caption         =   "Nueva Página"
         Height          =   375
         Left            =   4680
         TabIndex        =   29
         ToolTipText     =   "Nueva Pagina (Otra Fase del Mismo Evento)"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   32
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desencadenar"
      Height          =   735
      Index           =   18
      Left            =   2760
      TabIndex        =   24
      ToolTipText     =   "Consecuencia condicion desencadenante del Evento."
      Top             =   6840
      Width           =   3375
      Begin VB.ComboBox cmbTrigger 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":09F4
         Left            =   120
         List            =   "frmEditor_Events.frx":0A01
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1455
      Index           =   16
      Left            =   360
      TabIndex        =   20
      Top             =   6840
      Width           =   2295
      Begin VB.CheckBox chkShowName 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   337
         Top             =   960
         Width           =   1695
      End
      Begin VB.CheckBox chkWalkThrough 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pasar A Traves "
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   1695
      End
      Begin VB.CheckBox chkDirFix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Corregir Direccion"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1935
      End
      Begin VB.CheckBox chkWalkAnim 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Animacion Estatica"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Movimiento"
      Height          =   2175
      Index           =   15
      Left            =   2760
      TabIndex        =   13
      ToolTipText     =   "Si el Evento tiene movimiento podemos trazar su ruta o definirlo, ya que por ejemplo podria tratarse de un NPC evento."
      Top             =   3480
      Width           =   3375
      Begin VB.CommandButton cmdMoveRoute 
         Caption         =   "Ruta"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   98
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmbMoveFreq 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0A37
         Left            =   840
         List            =   "frmEditor_Events.frx":0A4A
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1680
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveSpeed 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0A74
         Left            =   840
         List            =   "frmEditor_Events.frx":0A8A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoveType 
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0AD8
         Left            =   840
         List            =   "frmEditor_Events.frx":0AE5
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Frec:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Veloc:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   390
         Width           =   1695
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagen"
      Height          =   3255
      Index           =   13
      Left            =   360
      TabIndex        =   11
      Top             =   3480
      Width           =   2295
      Begin VB.PictureBox picGraphic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   240
         ScaleHeight     =   193
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   12
         ToolTipText     =   "Imagen del Evento Actual."
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Condiciones"
      Height          =   2055
      Index           =   0
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Condiciones, Efectos que se Desencadenan, Variables y demas."
      Top             =   1320
      Width           =   5775
      Begin VB.ComboBox cmbPlayerVarCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0AFF
         Left            =   3720
         List            =   "frmEditor_Events.frx":0B15
         Style           =   2  'Dropdown List
         TabIndex        =   307
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbSelfSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0B68
         Left            =   4680
         List            =   "frmEditor_Events.frx":0B72
         Style           =   2  'Dropdown List
         TabIndex        =   306
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cmbPlayerSwitchCompare 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0B88
         Left            =   4680
         List            =   "frmEditor_Events.frx":0B92
         Style           =   2  'Dropdown List
         TabIndex        =   303
         Top             =   720
         Width           =   975
      End
      Begin VB.ComboBox cmbSelfSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0BA8
         Left            =   1920
         List            =   "frmEditor_Events.frx":0BBB
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CheckBox chkSelfSwitch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Propio Switch*"
         Height          =   255
         Left            =   120
         TabIndex        =   96
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkHasItem 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tiene Objeto     (En Inventario)"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.ComboBox cmbHasItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0BE4
         Left            =   1920
         List            =   "frmEditor_Events.frx":0BE6
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CheckBox chkPlayerSwitch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Switch Jugador"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayerSwitch 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0BE8
         Left            =   1920
         List            =   "frmEditor_Events.frx":0BEA
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkPlayerVar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Variable Jugador"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbPlayerVar 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Events.frx":0BEC
         Left            =   1920
         List            =   "frmEditor_Events.frx":0BEE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPlayerVariable 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "es"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   305
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "es"
         Height          =   255
         Index           =   3
         Left            =   4440
         TabIndex        =   304
         Top             =   795
         Width           =   255
      End
      Begin VB.Label lblRandomLabel 
         BackStyle       =   0  'Transparent
         Caption         =   "es"
         Height          =   255
         Index           =   5
         Left            =   3480
         TabIndex        =   6
         Top             =   340
         Width           =   255
      End
   End
   Begin VB.Frame fraRandom 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   735
      Index           =   9
      Left            =   6240
      TabIndex        =   178
      Top             =   7560
      Width           =   6255
      Begin VB.CommandButton cmdClearCommand 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   4680
         TabIndex        =   182
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdDeleteCommand 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3120
         TabIndex        =   181
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditCommand 
         Caption         =   "Editar"
         Height          =   375
         Left            =   1560
         TabIndex        =   180
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddCommand 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   179
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdLabel 
      Caption         =   "Variables/Switches"
      Height          =   375
      Left            =   120
      TabIndex        =   320
      Top             =   8520
      Width           =   2415
   End
   Begin VB.Frame fraDialogue 
      BackColor       =   &H00E0E0E0&
      Height          =   6975
      Left            =   6240
      TabIndex        =   73
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Variable Jugador"
         Height          =   2535
         Index           =   4
         Left            =   360
         TabIndex        =   82
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   3240
            TabIndex        =   346
            Text            =   "0"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   345
            Text            =   "0"
            Top             =   1590
            Width           =   855
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   1680
            TabIndex        =   344
            Text            =   "0"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtVariableData 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   1680
            TabIndex        =   343
            Text            =   "0"
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtVariableData 
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   342
            Text            =   "0"
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton optVariableAction 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Azar"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   341
            Top             =   1560
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sustraer"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   340
            Top             =   1320
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Agregar"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   339
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton optVariableAction 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   338
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdVariableCancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   86
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdVariableOK 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   85
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox cmbVariable 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0BF0
            Left            =   960
            List            =   "frmEditor_Events.frx":0BF2
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Alto:"
            Height          =   255
            Index           =   37
            Left            =   2760
            TabIndex        =   361
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Bajo:"
            Height          =   255
            Index           =   13
            Left            =   1440
            TabIndex        =   360
            Top             =   1590
            Width           =   495
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Variable:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agregue el texto"
         Height          =   4095
         Index           =   2
         Left            =   600
         TabIndex        =   216
         Top             =   720
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtAddText_Text 
            Height          =   1815
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   223
            Top             =   480
            Width           =   3855
         End
         Begin VB.HScrollBar scrlAddText_Colour 
            Height          =   255
            Left            =   120
            Max             =   18
            TabIndex        =   222
            Top             =   2640
            Width           =   3855
         End
         Begin VB.OptionButton optAddText_Player 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Jugador"
            Height          =   255
            Left            =   120
            TabIndex        =   221
            Top             =   3240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optAddText_Map 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mapa"
            Height          =   255
            Left            =   1080
            TabIndex        =   220
            Top             =   3240
            Width           =   855
         End
         Begin VB.OptionButton optAddText_Global 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Global"
            Height          =   255
            Left            =   1920
            TabIndex        =   219
            Top             =   3240
            Width           =   855
         End
         Begin VB.CommandButton cmdAddText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   218
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddText_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   217
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Texto:"
            Height          =   255
            Index           =   34
            Left            =   120
            TabIndex        =   226
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label lblAddText_Colour 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Color: Negro (No discrimines)"
            Height          =   255
            Left            =   120
            TabIndex        =   225
            Top             =   2400
            Width           =   3255
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Canal:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   224
            Top             =   3000
            Width           =   1575
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Burbuja Chat"
         Height          =   2775
         Index           =   3
         Left            =   600
         TabIndex        =   362
         Top             =   1680
         Visible         =   0   'False
         Width           =   5055
         Begin VB.TextBox txtChatbubbleText 
            Height          =   285
            Left            =   1680
            TabIndex        =   371
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox cmbChatBubbleTarget 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0BF4
            Left            =   1920
            List            =   "frmEditor_Events.frx":0BF6
            Style           =   2  'Dropdown List
            TabIndex        =   368
            Top             =   1560
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.OptionButton optChatBubbleTarget 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Evento"
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   367
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton optChatBubbleTarget 
            BackColor       =   &H00E0E0E0&
            Caption         =   "NPC"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   366
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optChatBubbleTarget 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Jugador"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   365
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdShowChatBubble_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2160
            TabIndex        =   364
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowChatBubble_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3600
            TabIndex        =   363
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Objetivo:"
            Height          =   255
            Index           =   39
            Left            =   120
            TabIndex        =   370
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Texto:"
            Height          =   255
            Index           =   38
            Left            =   120
            TabIndex        =   369
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Posibilidades"
         Height          =   4335
         Index           =   1
         Left            =   360
         TabIndex        =   188
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   4
            Left            =   2160
            TabIndex        =   199
            Text            =   "0"
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   197
            Text            =   "0"
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   2
            Left            =   2160
            TabIndex        =   195
            Text            =   "0"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.TextBox txtChoices 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   193
            Text            =   "0"
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CommandButton cmdChoices_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   191
            Top             =   3840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChoices_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   190
            Top             =   3840
            Width           =   1215
         End
         Begin VB.TextBox txtChoicePrompt 
            Height          =   1695
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   189
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opcion 4"
            Height          =   255
            Index           =   21
            Left            =   2160
            TabIndex        =   200
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opcion 3"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   198
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opcion 2"
            Height          =   255
            Index           =   19
            Left            =   2160
            TabIndex        =   196
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opcion 1"
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   194
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rapido:"
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   192
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Texto"
         Height          =   4095
         Index           =   0
         Left            =   600
         TabIndex        =   183
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtShowText 
            Height          =   2775
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   186
            Top             =   480
            Width           =   3855
         End
         Begin VB.CommandButton cmdShowText_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   185
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton cmdShowText_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   184
            Top             =   3600
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Texto:"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   187
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Script"
         Height          =   1575
         Index           =   29
         Left            =   360
         TabIndex        =   284
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
         Begin VB.HScrollBar scrlCustomScript 
            Height          =   255
            Left            =   1560
            Max             =   255
            Min             =   1
            TabIndex        =   288
            Top             =   360
            Value           =   1
            Width           =   2655
         End
         Begin VB.CommandButton cmdCustomScript_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   286
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdCustomScript_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3000
            TabIndex        =   285
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblCustomScript 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Caso: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   287
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acceso"
         Height          =   1575
         Index           =   28
         Left            =   240
         TabIndex        =   312
         Top             =   2760
         Visible         =   0   'False
         Width           =   4215
         Begin VB.ComboBox cmbSetAccess 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0BF8
            Left            =   960
            List            =   "frmEditor_Events.frx":0C0B
            Style           =   2  'Dropdown List
            TabIndex        =   315
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdSetAccess_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2880
            TabIndex        =   314
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetAccess_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   313
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Esperar"
         Height          =   1455
         Index           =   27
         Left            =   480
         TabIndex        =   419
         Top             =   3000
         Visible         =   0   'False
         Width           =   4335
         Begin VB.CommandButton cmdWait_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3000
            TabIndex        =   422
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdWait_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   421
            Top             =   840
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWaitAmount 
            Height          =   255
            Left            =   120
            Min             =   1
            TabIndex        =   420
            Top             =   480
            Value           =   1
            Width           =   4095
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "1000 Ms = 1 Segundo"
            Height          =   255
            Index           =   44
            Left            =   1920
            TabIndex        =   424
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblWaitAmount 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Esperar: 0 Ms"
            Height          =   255
            Left            =   120
            TabIndex        =   423
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sonido"
         Height          =   1575
         Index           =   26
         Left            =   240
         TabIndex        =   293
         Top             =   2760
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdPlaySound_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   296
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlaySound_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2880
            TabIndex        =   295
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbPlaySound 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0C57
            Left            =   960
            List            =   "frmEditor_Events.frx":0C59
            Style           =   2  'Dropdown List
            TabIndex        =   294
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BGM"
         Height          =   1575
         Index           =   25
         Left            =   240
         TabIndex        =   289
         Top             =   2640
         Visible         =   0   'False
         Width           =   4335
         Begin VB.ComboBox cmbPlayBGM 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0C5B
            Left            =   1080
            List            =   "frmEditor_Events.frx":0C5D
            Style           =   2  'Dropdown List
            TabIndex        =   292
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdPlayBGM_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3000
            TabIndex        =   291
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayBGM_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1560
            TabIndex        =   290
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Superposicion"
         Height          =   2055
         Index           =   24
         Left            =   480
         TabIndex        =   408
         Top             =   2520
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdMapTint_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   418
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CommandButton cmdMapTint_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   417
            Top             =   1560
            Width           =   1215
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   3
            Left            =   2280
            Max             =   255
            TabIndex        =   412
            Top             =   1200
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   411
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   1
            Left            =   2280
            Max             =   255
            TabIndex        =   410
            Top             =   480
            Width           =   855
         End
         Begin VB.HScrollBar scrlMapTintData 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   409
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label lblMapTintData 
            BackStyle       =   0  'Transparent
            Caption         =   "Opacidad: 0"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   416
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblMapTintData 
            BackStyle       =   0  'Transparent
            Caption         =   "Rojo: 0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   415
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblMapTintData 
            BackStyle       =   0  'Transparent
            Caption         =   "Verde: 0"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   414
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblMapTintData 
            BackStyle       =   0  'Transparent
            Caption         =   "Azul: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   413
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clima"
         Height          =   1935
         Index           =   23
         Left            =   480
         TabIndex        =   401
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSetWeather_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   407
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetWeather_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   406
            Top             =   1440
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWeatherIntensity 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   403
            Top             =   1080
            Width           =   1815
         End
         Begin VB.ComboBox CmbWeather 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0C5F
            Left            =   120
            List            =   "frmEditor_Events.frx":0C75
            Style           =   2  'Dropdown List
            TabIndex        =   402
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lblWeatherIntensity 
            BackStyle       =   0  'Transparent
            Caption         =   "Intensidad: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   405
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clima:"
            Height          =   195
            Index           =   43
            Left            =   120
            TabIndex        =   404
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Niebla"
         Height          =   2415
         Index           =   22
         Left            =   480
         TabIndex        =   392
         Top             =   2280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSetFog_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   400
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdSetFog_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   399
            Top             =   1920
            Width           =   1215
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   1
            Left            =   120
            Max             =   255
            TabIndex        =   395
            Top             =   1050
            Width           =   1575
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   0
            Left            =   120
            Max             =   255
            TabIndex        =   394
            Top             =   480
            Width           =   1575
         End
         Begin VB.HScrollBar ScrlFogData 
            Height          =   255
            Index           =   2
            Left            =   120
            Max             =   255
            TabIndex        =   393
            Top             =   1620
            Width           =   1575
         End
         Begin VB.Label lblFogData 
            BackStyle       =   0  'Transparent
            Caption         =   "Niebla Vel: 0"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   398
            Top             =   810
            Width           =   1815
         End
         Begin VB.Label lblFogData 
            BackStyle       =   0  'Transparent
            Caption         =   "Niebla: No"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   397
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblFogData 
            BackStyle       =   0  'Transparent
            Caption         =   "Niebla Opac: 0"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   396
            Top             =   1380
            Width           =   1815
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tienda"
         Height          =   1575
         Index           =   21
         Left            =   600
         TabIndex        =   316
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
         Begin VB.CommandButton cmdOpenShop_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   319
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdOpenShop_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2880
            TabIndex        =   318
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbOpenShop 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0CB1
            Left            =   960
            List            =   "frmEditor_Events.frx":0CC4
            Style           =   2  'Dropdown List
            TabIndex        =   317
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Spawn NPC"
         Height          =   1695
         Index           =   19
         Left            =   360
         TabIndex        =   387
         Top             =   2280
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbSpawnNPC 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D10
            Left            =   120
            List            =   "frmEditor_Events.frx":0D12
            Style           =   2  'Dropdown List
            TabIndex        =   391
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdSpawnNpc_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   389
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSpawnNpc_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2640
            TabIndex        =   388
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "NPC:"
            Height          =   255
            Index           =   42
            Left            =   120
            TabIndex        =   390
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reproducir Anim"
         Height          =   2775
         Index           =   20
         Left            =   240
         TabIndex        =   270
         Top             =   1800
         Visible         =   0   'False
         Width           =   5055
         Begin VB.ComboBox cmbPlayAnim 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D14
            Left            =   1680
            List            =   "frmEditor_Events.frx":0D16
            Style           =   2  'Dropdown List
            TabIndex        =   283
            Top             =   300
            Width           =   3135
         End
         Begin VB.HScrollBar scrlPlayAnimTileY 
            Height          =   255
            Left            =   1920
            TabIndex        =   281
            Top             =   1800
            Width           =   2895
         End
         Begin VB.HScrollBar scrlPlayAnimTileX 
            Height          =   255
            Left            =   1920
            TabIndex        =   280
            Top             =   1455
            Width           =   2895
         End
         Begin VB.CommandButton cmdPlayAnim_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3600
            TabIndex        =   276
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayAnim_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2160
            TabIndex        =   275
            Top             =   2280
            Width           =   1215
         End
         Begin VB.OptionButton optPlayAnimPlayer 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Jugador"
            Height          =   255
            Left            =   120
            TabIndex        =   274
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optPlayAnimEvent 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Evento"
            Height          =   255
            Left            =   1920
            TabIndex        =   273
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton optPlayAnimTile 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tile"
            Height          =   255
            Left            =   3720
            TabIndex        =   272
            Top             =   1080
            Width           =   975
         End
         Begin VB.ComboBox cmbPlayAnimEvent 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D18
            Left            =   1920
            List            =   "frmEditor_Events.frx":0D1A
            Style           =   2  'Dropdown List
            TabIndex        =   271
            Top             =   1440
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Animacion"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   282
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimY 
            BackStyle       =   0  'Transparent
            Caption         =   "Tile Y:"
            Height          =   255
            Left            =   240
            TabIndex        =   279
            Top             =   1800
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblPlayAnimX 
            BackStyle       =   0  'Transparent
            Caption         =   "Tile X:"
            Height          =   255
            Left            =   240
            TabIndex        =   278
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Objetivo:"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   277
            Top             =   840
            Width           =   1335
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Transportar"
         Height          =   3015
         Index           =   18
         Left            =   240
         TabIndex        =   87
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbWarpPlayerDir 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D1C
            Left            =   120
            List            =   "frmEditor_Events.frx":0D2F
            Style           =   2  'Dropdown List
            TabIndex        =   301
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CommandButton cmdWPCancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   95
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton cmdWPOkay 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   94
            Top             =   2520
            Width           =   1215
         End
         Begin VB.HScrollBar scrlWPY 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   93
            Top             =   1680
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPX 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   91
            Top             =   1080
            Width           =   3855
         End
         Begin VB.HScrollBar scrlWPMap 
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblWPY 
            BackStyle       =   0  'Transparent
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1440
            Width           =   3135
         End
         Begin VB.Label lblWPX 
            BackStyle       =   0  'Transparent
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label lblWPMap 
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dar Experiencia"
         Height          =   2415
         Index           =   17
         Left            =   240
         TabIndex        =   372
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbSkilling 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D6E
            Left            =   120
            List            =   "frmEditor_Events.frx":0D70
            Style           =   2  'Dropdown List
            TabIndex        =   428
            Top             =   600
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.OptionButton opSkilling 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Habilidad Exp"
            Height          =   255
            Left            =   2400
            TabIndex        =   427
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton opMine 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Combate Exp"
            Height          =   255
            Left            =   120
            TabIndex        =   426
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.HScrollBar scrlGiveExp 
            Height          =   255
            Left            =   120
            Max             =   32000
            TabIndex        =   375
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CommandButton cmdGiveExp_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2640
            TabIndex        =   374
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton cmdGiveExp_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   373
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblGiveExp 
            BackStyle       =   0  'Transparent
            Caption         =   "Dar Exp: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   376
            Top             =   1080
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PK"
         Height          =   1455
         Index           =   16
         Left            =   240
         TabIndex        =   263
         Top             =   3000
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangePK_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   267
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangePK_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   266
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton optChangePKYes 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Si"
            Height          =   255
            Left            =   240
            TabIndex        =   265
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optChangePKNo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "No"
            Height          =   255
            Left            =   1920
            TabIndex        =   264
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Genero"
         Height          =   1455
         Index           =   15
         Left            =   240
         TabIndex        =   258
         Top             =   2760
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSexFemale 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mujer"
            Height          =   255
            Left            =   1920
            TabIndex        =   262
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optChangeSexMale 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hombre"
            Height          =   255
            Left            =   240
            TabIndex        =   261
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdChangeSex_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   260
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSex_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   259
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Sprite"
         Height          =   1695
         Index           =   14
         Left            =   240
         TabIndex        =   253
         Top             =   2640
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeSprite 
            Height          =   255
            Left            =   1200
            Max             =   100
            Min             =   1
            TabIndex        =   257
            Top             =   360
            Value           =   1
            Width           =   2535
         End
         Begin VB.CommandButton cmdChangeSprite_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   255
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSprite_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   254
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblChangeSprite 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sprite: 1"
            Height          =   255
            Left            =   120
            TabIndex        =   256
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Clase"
         Height          =   1695
         Index           =   13
         Left            =   480
         TabIndex        =   248
         Top             =   2520
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdChangeClass_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   251
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeClass_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   250
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cmbChangeClass 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D72
            Left            =   120
            List            =   "frmEditor_Events.frx":0D74
            Style           =   2  'Dropdown List
            TabIndex        =   249
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clase:"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   252
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Skills"
         Height          =   2175
         Index           =   12
         Left            =   360
         TabIndex        =   241
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
         Begin VB.OptionButton optChangeSkillsRemove 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            Height          =   255
            Left            =   1800
            TabIndex        =   247
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optChangeSkillsAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Instruir"
            Height          =   255
            Left            =   120
            TabIndex        =   246
            Top             =   960
            Width           =   1455
         End
         Begin VB.ComboBox cmbChangeSkills 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D76
            Left            =   120
            List            =   "frmEditor_Events.frx":0D78
            Style           =   2  'Dropdown List
            TabIndex        =   245
            Top             =   480
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeSkills_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   243
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeSkills_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   242
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Habilidad"
            Height          =   255
            Index           =   28
            Left            =   120
            TabIndex        =   244
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Nivel"
         Height          =   1815
         Index           =   11
         Left            =   240
         TabIndex        =   236
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
         Begin VB.HScrollBar scrlChangeLevel 
            Height          =   255
            Left            =   120
            TabIndex        =   240
            Top             =   600
            Width           =   3615
         End
         Begin VB.CommandButton cmdChangeLevel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   238
            Top             =   1080
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeLevel_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2520
            TabIndex        =   237
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblChangeLevel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   239
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar Objetos"
         Height          =   2415
         Index           =   10
         Left            =   240
         TabIndex        =   227
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbChangeItemIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D7A
            Left            =   120
            List            =   "frmEditor_Events.frx":0D7C
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtChangeItemsAmount 
            Height          =   375
            Left            =   120
            TabIndex        =   234
            Text            =   "0"
            Top             =   1320
            Width           =   3735
         End
         Begin VB.CommandButton cmdChangeItems_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2640
            TabIndex        =   232
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton cmdChangeItems_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1200
            TabIndex        =   231
            Top             =   1800
            Width           =   1215
         End
         Begin VB.OptionButton optChangeItemRemove 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Quitar"
            Height          =   255
            Left            =   2640
            TabIndex        =   230
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optChangeItemAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dar"
            Height          =   255
            Left            =   1680
            TabIndex        =   229
            Top             =   960
            Width           =   735
         End
         Begin VB.OptionButton optChangeItemSet 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cantidad"
            Height          =   255
            Left            =   120
            TabIndex        =   228
            Top             =   960
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto Index:"
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   233
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IrA Etiquetal"
         Height          =   1695
         Index           =   9
         Left            =   480
         TabIndex        =   382
         Top             =   1680
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdGotoLabel_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2640
            TabIndex        =   385
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdGotoLabel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   384
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtGotoLabel 
            Height          =   375
            Left            =   120
            TabIndex        =   383
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Etiqueta:"
            Height          =   255
            Index           =   41
            Left            =   120
            TabIndex        =   386
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Crear Etiqueta"
         Height          =   1695
         Index           =   8
         Left            =   360
         TabIndex        =   377
         Top             =   1800
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtLabelName 
            Height          =   375
            Left            =   120
            TabIndex        =   381
            Top             =   480
            Width           =   3855
         End
         Begin VB.CommandButton cmdCreateLabel_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1320
            TabIndex        =   379
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdCreateLabel_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2640
            TabIndex        =   378
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Etiqueta:"
            Height          =   255
            Index           =   40
            Left            =   120
            TabIndex        =   380
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Propio Switch"
         Height          =   1695
         Index           =   6
         Left            =   360
         TabIndex        =   208
         Top             =   1320
         Visible         =   0   'False
         Width           =   4095
         Begin VB.CommandButton cmdSelfSwitch_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   212
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdSelfSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   211
            Top             =   1200
            Width           =   1215
         End
         Begin VB.ComboBox cmbSetSelfSwitch 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D7E
            Left            =   1440
            List            =   "frmEditor_Events.frx":0D8E
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   360
            Width           =   2535
         End
         Begin VB.ComboBox cmbSetSelfSwitchTo 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0D9E
            Left            =   960
            List            =   "frmEditor_Events.frx":0DA8
            Style           =   2  'Dropdown List
            TabIndex        =   209
            Top             =   800
            Width           =   3015
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Modif a:"
            Height          =   255
            Index           =   26
            Left            =   120
            TabIndex        =   214
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "Propio Switch:"
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   213
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Condición"
         Height          =   5775
         Index           =   7
         Left            =   240
         TabIndex        =   158
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
         Begin VB.ComboBox cmbCondition_Status 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0DC0
            Left            =   4680
            List            =   "frmEditor_Events.frx":0DCD
            Style           =   2  'Dropdown List
            TabIndex        =   436
            Top             =   4680
            Width           =   1335
         End
         Begin VB.ComboBox cmbCondition_Quest 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0DF0
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E00
            Style           =   2  'Dropdown List
            TabIndex        =   434
            Top             =   4680
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Misión Jugador"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   433
            Top             =   4680
            Width           =   1575
         End
         Begin VB.TextBox txtCondition_SkillLvlReq 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4080
            TabIndex        =   432
            Text            =   "0"
            Top             =   4200
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_SkillReq 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E10
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E20
            Style           =   2  'Dropdown List
            TabIndex        =   430
            Top             =   4200
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel Habilidad"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   429
            Top             =   4200
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_itemAmount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3840
            TabIndex        =   425
            Text            =   "0"
            Top             =   1800
            Width           =   855
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Propio Switch"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   299
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E30
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E40
            Style           =   2  'Dropdown List
            TabIndex        =   298
            Top             =   3720
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_SelfSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E50
            Left            =   3960
            List            =   "frmEditor_Events.frx":0E5A
            Style           =   2  'Dropdown List
            TabIndex        =   297
            Top             =   3720
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_LearntSkill 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E70
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E72
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   2760
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_ClassIs 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E74
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E76
            Style           =   2  'Dropdown List
            TabIndex        =   176
            Top             =   2280
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondition_HasItem 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E78
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E7A
            Style           =   2  'Dropdown List
            TabIndex        =   175
            Top             =   1800
            Width           =   1695
         End
         Begin VB.ComboBox cmbCondtion_PlayerSwitchCondition 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E7C
            Left            =   3960
            List            =   "frmEditor_Events.frx":0E86
            Style           =   2  'Dropdown List
            TabIndex        =   174
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondition_PlayerSwitch 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0E9C
            Left            =   1920
            List            =   "frmEditor_Events.frx":0E9E
            Style           =   2  'Dropdown List
            TabIndex        =   173
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_LevelAmount 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3480
            TabIndex        =   172
            Text            =   "0"
            Top             =   3240
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_LevelCompare 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0EA0
            Left            =   1440
            List            =   "frmEditor_Events.frx":0EB6
            Style           =   2  'Dropdown List
            TabIndex        =   171
            Top             =   3240
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Nivel"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   170
            Top             =   3240
            Width           =   975
         End
         Begin VB.ComboBox cmbCondition_PlayerVarCompare 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0F09
            Left            =   1920
            List            =   "frmEditor_Events.frx":0F1F
            Style           =   2  'Dropdown List
            TabIndex        =   168
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtCondition_PlayerVarCondition 
            Height          =   285
            Left            =   3840
            TabIndex        =   167
            Text            =   "0"
            Top             =   840
            Width           =   855
         End
         Begin VB.ComboBox cmbCondition_PlayerVarIndex 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0F72
            Left            =   1920
            List            =   "frmEditor_Events.frx":0F74
            Style           =   2  'Dropdown List
            TabIndex        =   166
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Habilidad"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   165
            Top             =   2760
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Clase"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   164
            Top             =   2280
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Tiene Objeto"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   163
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Switch Jugador"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   162
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optCondition_Index 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Variable Jugador"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   161
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton cmdCondition_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   3360
            TabIndex        =   160
            Top             =   5160
            Width           =   1215
         End
         Begin VB.CommandButton cmdCondition_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   4680
            TabIndex        =   159
            Top             =   5160
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "   Estado"
            Height          =   255
            Index           =   46
            Left            =   3600
            TabIndex        =   435
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   ">="
            Height          =   255
            Index           =   45
            Left            =   3720
            TabIndex        =   431
            Top             =   4200
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "es"
            Height          =   255
            Index           =   35
            Left            =   3720
            TabIndex        =   300
            Top             =   3720
            Width           =   255
         End
         Begin VB.Label lblRandomLabel 
            BackStyle       =   0  'Transparent
            Caption         =   "es"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   169
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.Frame fraCommand 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Switch Jugador"
         Height          =   1695
         Index           =   5
         Left            =   240
         TabIndex        =   201
         Top             =   1560
         Visible         =   0   'False
         Width           =   4095
         Begin VB.ComboBox cmbPlayerSwitchSet 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0F76
            Left            =   1200
            List            =   "frmEditor_Events.frx":0F80
            Style           =   2  'Dropdown List
            TabIndex        =   207
            Top             =   800
            Width           =   2775
         End
         Begin VB.ComboBox cmbSwitch 
            Height          =   315
            ItemData        =   "frmEditor_Events.frx":0F96
            Left            =   1200
            List            =   "frmEditor_Events.frx":0F98
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   360
            Width           =   2775
         End
         Begin VB.CommandButton cmbPlayerSwitch_Ok 
            Caption         =   "Ok"
            Height          =   375
            Left            =   1440
            TabIndex        =   203
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdPlayerSwitch_Cancel 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   2760
            TabIndex        =   202
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Switch:"
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   206
            Top             =   360
            Width           =   3855
         End
         Begin VB.Label lblRandomLabel 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Cambiar a:"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   205
            Top             =   840
            Width           =   1815
         End
      End
   End
   Begin VB.Frame fraCommands 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Comandos"
      Height          =   6975
      Left            =   6240
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   6375
      Begin VB.CommandButton cmdCancelCommand 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4560
         TabIndex        =   72
         Top             =   6360
         Width           =   1455
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6135
         Index           =   1
         Left            =   240
         ScaleHeight     =   6135
         ScaleWidth      =   5775
         TabIndex        =   39
         Top             =   600
         Width           =   5775
         Begin VB.Frame fraRandom 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Control Jugador"
            Height          =   5535
            Index           =   3
            Left            =   3000
            TabIndex        =   52
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Dar EXP"
               Height          =   375
               Index           =   21
               Left            =   120
               TabIndex        =   336
               Top             =   5040
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar PK"
               Height          =   375
               Index           =   20
               Left            =   120
               TabIndex        =   215
               Top             =   4560
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Sexo"
               Height          =   375
               Index           =   19
               Left            =   120
               TabIndex        =   61
               Top             =   4080
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Sprite"
               Height          =   375
               Index           =   18
               Left            =   120
               TabIndex        =   60
               Top             =   3600
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Clase"
               Height          =   375
               Index           =   17
               Left            =   120
               TabIndex        =   59
               Top             =   3120
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Nivel"
               Height          =   375
               Index           =   15
               Left            =   120
               TabIndex        =   57
               Top             =   2160
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Subir Nivel"
               Height          =   375
               Index           =   14
               Left            =   120
               TabIndex        =   56
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Regen MP"
               Height          =   375
               Index           =   13
               Left            =   120
               TabIndex        =   55
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Regen HP"
               Height          =   375
               Index           =   12
               Left            =   120
               TabIndex        =   54
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Objetos"
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   53
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cambiar Habilidad"
               Height          =   375
               Index           =   16
               Left            =   120
               TabIndex        =   58
               Top             =   2640
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Control de Flujo"
            Height          =   2175
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   3720
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Ir a Etiqueta"
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   352
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Etiqueta"
               Height          =   375
               Index           =   9
               Left            =   120
               TabIndex        =   351
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Cadena Condicional"
               Height          =   375
               Index           =   7
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Terminar Evento"
               Height          =   375
               Index           =   8
               Left            =   120
               TabIndex        =   50
               Top             =   720
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Progreso Evento"
            Height          =   1695
            Index           =   1
            Left            =   0
            TabIndex        =   45
            Top             =   2160
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Switch Propio"
               Height          =   375
               Index           =   6
               Left            =   120
               TabIndex        =   48
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Switch Jugador"
               Height          =   375
               Index           =   5
               Left            =   120
               TabIndex        =   47
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Variable Jugador"
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mensaje"
            Height          =   2175
            Index           =   21
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Mostrar Opciones"
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Burbuja de Chat"
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   347
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Mostrar Texto"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Texto de ChatBox"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   42
               Top             =   1200
               Width           =   2535
            End
         End
      End
      Begin VB.PictureBox picCommands 
         BorderStyle     =   0  'None
         Height          =   6015
         Index           =   2
         Left            =   240
         ScaleHeight     =   6015
         ScaleWidth      =   5775
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame fraRandom 
            Caption         =   "Funciones Mapa"
            Height          =   1695
            Index           =   12
            Left            =   3000
            TabIndex        =   356
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Niebla..."
               Height          =   375
               Index           =   31
               Left            =   120
               TabIndex        =   359
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Clima..."
               Height          =   375
               Index           =   32
               Left            =   120
               TabIndex        =   358
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Matizado de Mapa"
               Height          =   375
               Index           =   33
               Left            =   120
               TabIndex        =   357
               Top             =   1200
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Corte Escena Opciones"
            Height          =   1695
            Index           =   11
            Left            =   0
            TabIndex        =   349
            Top             =   3840
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Flash Blanco"
               Height          =   375
               Index           =   30
               Left            =   120
               TabIndex        =   355
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Fade Out"
               Height          =   375
               Index           =   29
               Left            =   120
               TabIndex        =   354
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Fade In"
               Height          =   375
               Index           =   28
               Left            =   120
               TabIndex        =   350
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Tienda y Banco"
            Height          =   1215
            Index           =   6
            Left            =   0
            TabIndex        =   308
            Top             =   2520
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Abrir Tienda"
               Height          =   375
               Index           =   27
               Left            =   120
               TabIndex        =   310
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Abrir Banco"
               Height          =   375
               Index           =   26
               Left            =   120
               TabIndex        =   309
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Etc..."
            Height          =   1695
            Index           =   8
            Left            =   3000
            TabIndex        =   268
            Top             =   3840
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Esperar..."
               Height          =   375
               Index           =   38
               Left            =   120
               TabIndex        =   348
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Privilegios..."
               Height          =   375
               Index           =   39
               Left            =   120
               TabIndex        =   311
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Script"
               Height          =   375
               Index           =   40
               Left            =   120
               TabIndex        =   269
               Top             =   1200
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "BGM y SFX"
            Height          =   2175
            Index           =   7
            Left            =   3000
            TabIndex        =   67
            Top             =   1680
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Frenar SFX"
               Height          =   375
               Index           =   37
               Left            =   120
               TabIndex        =   71
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "SFX"
               Height          =   375
               Index           =   36
               Left            =   120
               TabIndex        =   70
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Apagar BGM"
               Height          =   375
               Index           =   35
               Left            =   120
               TabIndex        =   69
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "BGM"
               Height          =   375
               Index           =   34
               Left            =   120
               TabIndex        =   68
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Animacion"
            Height          =   735
            Index           =   5
            Left            =   0
            TabIndex        =   65
            Top             =   1680
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Reprod Animacion"
               Height          =   375
               Index           =   25
               Left            =   120
               TabIndex        =   66
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame fraRandom 
            Caption         =   "Movimiento"
            Height          =   1695
            Index           =   4
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   2775
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Spawnear"
               Height          =   375
               Index           =   24
               Left            =   120
               TabIndex        =   353
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Modificar Ruta"
               Height          =   375
               Index           =   23
               Left            =   120
               TabIndex        =   64
               Top             =   720
               Width           =   2535
            End
            Begin VB.CommandButton cmdCommands 
               Caption         =   "Transportar"
               Height          =   375
               Index           =   22
               Left            =   120
               TabIndex        =   63
               Top             =   240
               Width           =   2535
            End
         End
      End
      Begin MSComctlLib.TabStrip tabCommands 
         Height          =   6615
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   11668
         MultiRow        =   -1  'True
         TabMinWidth     =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "1"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "2"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.ListBox lstCommands 
      Height          =   6105
      ItemData        =   "frmEditor_Events.frx":0F9A
      Left            =   6240
      List            =   "frmEditor_Events.frx":0F9C
      TabIndex        =   1
      Top             =   1440
      Width           =   6255
   End
   Begin MSComctlLib.TabStrip tabPages 
      Height          =   7455
      Left            =   240
      TabIndex        =   34
      Top             =   960
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   13150
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      TabMinWidth     =   529
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "1"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblRandomLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* Propio Switch es Global y se reinicia al arrancar el Server."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   302
      Top             =   8520
      Width           =   4935
   End
   Begin VB.Label lblRandomLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "www.easee.es                                "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   14
      Left            =   2640
      TabIndex        =   101
      Top             =   8700
      Width           =   6975
   End
   Begin VB.Label lblRandomLabel 
      Caption         =   "Lista Comandos:"
      Height          =   255
      Index           =   9
      Left            =   6240
      TabIndex        =   0
      Top             =   1560
      Width           =   6255
   End
End
Attribute VB_Name = "frmEditor_Events"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private copyPage As EventPageRec

Private Sub chkDirFix_Click()
    tmpEvent.Pages(curPageNum).DirFix = chkDirFix.Value
End Sub

Private Sub chkGlobal_Click()
    tmpEvent.Global = chkGlobal.Value
End Sub

Private Sub chkHasItem_Click()
    tmpEvent.Pages(curPageNum).chkHasItem = chkHasItem.Value
    If chkHasItem.Value = 0 Then cmbHasItem.Enabled = False Else cmbHasItem.Enabled = True
End Sub

Private Sub chkIgnoreMove_Click()
    tmpEvent.Pages(curPageNum).IgnoreMoveRoute = chkIgnoreMove.Value
End Sub

Private Sub chkPlayerSwitch_Click()
    tmpEvent.Pages(curPageNum).chkSwitch = chkPlayerSwitch.Value
    If chkPlayerSwitch.Value = 0 Then
        cmbPlayerSwitch.Enabled = False
        cmbPlayerSwitchCompare.Enabled = False
    Else
        cmbPlayerSwitch.Enabled = True
        cmbPlayerSwitchCompare.Enabled = True
    End If
End Sub

Private Sub chkPlayerVar_Click()
    tmpEvent.Pages(curPageNum).chkVariable = chkPlayerVar.Value
    If chkPlayerVar.Value = 0 Then
        cmbPlayerVar.Enabled = False
        txtPlayerVariable.Enabled = False
        cmbPlayerVarCompare.Enabled = False
    Else
        cmbPlayerVar.Enabled = True
        txtPlayerVariable.Enabled = True
        cmbPlayerVarCompare.Enabled = True
    End If
End Sub

Private Sub chkRepeatRoute_Click()
    tmpEvent.Pages(curPageNum).RepeatMoveRoute = chkRepeatRoute.Value
End Sub

Private Sub chkSelfSwitch_Click()
    tmpEvent.Pages(curPageNum).chkSelfSwitch = chkSelfSwitch.Value
    If chkSelfSwitch.Value = 0 Then
        cmbSelfSwitch.Enabled = False
        cmbSelfSwitchCompare.Enabled = False
    Else
        cmbSelfSwitch.Enabled = True
        cmbSelfSwitchCompare.Enabled = True
    End If
End Sub

Private Sub chkShowName_Click()
    tmpEvent.Pages(curPageNum).ShowName = chkShowName.Value
End Sub

Private Sub chkWalkAnim_Click()
    tmpEvent.Pages(curPageNum).WalkAnim = chkWalkAnim.Value
End Sub

Private Sub chkWalkThrough_Click()
    tmpEvent.Pages(curPageNum).Walkthrough = chkWalkThrough.Value
End Sub

Private Sub cmbGraphic_Click()
    If cmbGraphic.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).GraphicType = cmbGraphic.ListIndex
    ' set the max on the scrollbar
    Select Case cmbGraphic.ListIndex
        Case 0 ' None
            scrlGraphic.Value = 1
            scrlGraphic.Enabled = False
        Case 1 ' character
            scrlGraphic.max = NumCharacters
            scrlGraphic.Enabled = True
        Case 2 ' Tileset
            scrlGraphic.max = NumTileSets
            scrlGraphic.Enabled = True
    End Select
    
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Numero: Ninguno"
    Else
        lblGraphic.Caption = "Numero: " & scrlGraphic.Value
    End If
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    End If
End Sub

Private Sub cmbHasItem_Click()
    If cmbHasItem.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).HasItemIndex = cmbHasItem.ListIndex
    tmpEvent.Pages(curPageNum).HasItemAmount = txtCondition_itemAmount.text
End Sub

Private Sub cmbLabel_Ok_Click()
    fraLabeling.Visible = False
    lstCommands.Visible = True
    SendSwitchesAndVariables
End Sub

Private Sub cmbMoveFreq_Click()
    If cmbMoveFreq.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveFreq = cmbMoveFreq.ListIndex
End Sub

Private Sub cmbMoveSpeed_Click()
    If cmbMoveSpeed.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveSpeed = cmbMoveSpeed.ListIndex
End Sub

Private Sub cmbMoveType_Click()
    If cmbMoveType.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).MoveType = cmbMoveType.ListIndex
    If cmbMoveType.ListIndex = 2 Then
        cmdMoveRoute.Enabled = True
    Else
        cmdMoveRoute.Enabled = False
    End If
End Sub

Private Sub cmbPlayerSwitch_Click()
    If cmbPlayerSwitch.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SwitchIndex = cmbPlayerSwitch.ListIndex
End Sub

Private Sub cmbPlayerSwitch_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerSwitch
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(5).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmbPlayerSwitchCompare_Click()
    If cmbPlayerSwitchCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SwitchCompare = cmbPlayerSwitchCompare.ListIndex
End Sub

Private Sub cmbPlayerVar_Click()
    If cmbPlayerVar.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).VariableIndex = cmbPlayerVar.ListIndex
End Sub

Private Sub cmbPlayerVarCompare_Click()
    If cmbPlayerVarCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).VariableCompare = cmbPlayerVarCompare.ListIndex
End Sub

Private Sub cmbPositioning_Click()
    If cmbPositioning.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).Position = cmbPositioning.ListIndex
End Sub

Private Sub cmbSelfSwitch_Click()
    If cmbSelfSwitch.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SelfSwitchIndex = cmbSelfSwitch.ListIndex
End Sub

Private Sub cmbSelfSwitchCompare_Click()
    If cmbSelfSwitchCompare.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).SelfSwitchCompare = cmbSelfSwitchCompare.ListIndex
End Sub

Private Sub cmbTrigger_Click()
    If cmbTrigger.ListIndex = -1 Then Exit Sub
    tmpEvent.Pages(curPageNum).Trigger = cmbTrigger.ListIndex
End Sub

Private Sub cmdAddCommand_Click()
    If lstCommands.ListIndex > -1 Then
        isEdit = False
        tabCommands.SelectedItem = tabCommands.Tabs(1)
        fraCommands.Visible = True
        picCommands(1).Visible = True
        picCommands(2).Visible = False
    End If
End Sub

Private Sub cmdAddMoveRoute_Click(Index As Integer)
    If Index = 42 Then
        fraGraphic.Width = 841
        fraGraphic.Height = 585
        fraGraphic.Visible = True
        GraphicSelType = 1
    Else
        AddMoveRouteCommand Index
    End If
End Sub

Private Sub cmdAddText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(2).Visible = False
End Sub

Private Sub cmdAddText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evAddText
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(2).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelCommand_Click()
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeClass_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(13).Visible = False
End Sub

Private Sub cmdChangeClass_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeClass
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(13).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeItems_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(10).Visible = False
End Sub

Private Sub cmdChangeItems_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeItems
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommands.Visible = False
    fraCommand(10).Visible = False
End Sub

Private Sub cmdChangeLevel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(11).Visible = False
End Sub

Private Sub cmdChangeLevel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeLevel
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(11).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangePK_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(16).Visible = False
End Sub

Private Sub cmdChangePK_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangePK
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(16).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSex_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(15).Visible = False
End Sub

Private Sub cmdChangeSex_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSex
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(15).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSkills_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(12).Visible = False
End Sub

Private Sub cmdChangeSkills_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSkills
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(12).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChangeSprite_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(14).Visible = False
End Sub

Private Sub cmdChangeSprite_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evChangeSprite
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(14).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdChoices_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(1).Visible = False
End Sub

Private Sub cmdChoices_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChoices
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(1).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdClearCommand_Click()
    If Msgbox("Estas seguro?", vbYesNo, "Eliminar Comandos Eventos?") = vbYes Then
        ClearEventCommands
    End If
End Sub

Private Sub cmdClearPage_Click()
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
End Sub

Private Sub cmdCommands_Click(Index As Integer)
Dim I As Long, X As Long
    Select Case Index
        Case 0
            txtShowText.text = vbNullString
            fraDialogue.Visible = True
            fraCommand(0).Visible = True
            fraCommands.Visible = False
        Case 1
            txtChoicePrompt.text = vbNullString
            txtChoices(1).text = vbNullString
            txtChoices(2).text = vbNullString
            txtChoices(3).text = vbNullString
            txtChoices(4).text = vbNullString
            fraDialogue.Visible = True
            fraCommand(1).Visible = True
            fraCommands.Visible = False
        Case 2
            txtAddText_Text.text = vbNullString
            scrlAddText_Colour.Value = 0
            optAddText_Player.Value = True
            fraDialogue.Visible = True
            fraCommand(2).Visible = True
            fraCommands.Visible = False
        Case 3
            txtChatbubbleText.text = ""
            optChatBubbleTarget(0).Value = True
            cmbChatBubbleTarget.Visible = False
            fraDialogue.Visible = True
            fraCommand(3).Visible = True
            fraCommands.Visible = False
        Case 4
            For I = 0 To 4
                txtVariableData(I).text = 0
            Next
            cmbVariable.ListIndex = 0
            optVariableAction(0).Value = True
            fraDialogue.Visible = True
            fraCommand(4).Visible = True
            fraCommands.Visible = False
        Case 5
            cmbPlayerSwitchSet.ListIndex = 0
            cmbSwitch.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(5).Visible = True
            fraCommands.Visible = False
        Case 6
            cmbSetSelfSwitch.ListIndex = 0
            cmbSetSelfSwitchTo.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(6).Visible = True
            fraCommands.Visible = False
        Case 7
            fraDialogue.Visible = True
            fraCommand(7).Visible = True
            optCondition_Index(0).Value = True
            ClearConditionFrame
            cmbCondition_PlayerVarIndex.Enabled = True
            cmbCondition_PlayerVarCompare.Enabled = True
            txtCondition_PlayerVarCondition.Enabled = True
            fraCommands.Visible = False
        Case 8
            AddCommand EventType.evExitProcess
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 9
            txtLabelName.text = ""
            fraCommand(8).Visible = True
            fraCommands.Visible = False
            fraDialogue.Visible = True
        Case 10
            txtGotoLabel.text = ""
            fraCommand(9).Visible = True
            fraCommands.Visible = False
            fraDialogue.Visible = True
        Case 11
            cmbChangeItemIndex.ListIndex = 0
            optChangeItemSet.Value = True
            txtChangeItemsAmount.text = "0"
            fraDialogue.Visible = True
            fraCommand(10).Visible = True
            fraCommands.Visible = False
        Case 12
            AddCommand EventType.evRestoreHP
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 13
            AddCommand EventType.evRestoreMP
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 14
            AddCommand EventType.evLevelUp
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 15
            scrlChangeLevel.Value = 1
            lblChangeLevel.Caption = "Nivel: 1"
            fraDialogue.Visible = True
            fraCommand(11).Visible = True
            fraCommands.Visible = False
        Case 16
            cmbChangeSkills.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(12).Visible = True
            fraCommands.Visible = False
        Case 17
            If Max_Classes > 0 Then
                If cmbChangeClass.ListCount = 0 Then
                cmbChangeClass.Clear
                For I = 1 To Max_Classes
                    cmbChangeClass.AddItem Trim$(Class(I).name)
                Next
                cmbChangeClass.ListIndex = 0
                End If
            End If
            fraDialogue.Visible = True
            fraCommand(13).Visible = True
            fraCommands.Visible = False
        Case 18
            scrlChangeSprite.Value = 1
            lblChangeSprite.Caption = "Sprite: 1"
            fraDialogue.Visible = True
            fraCommand(14).Visible = True
            fraCommands.Visible = False
        Case 19
            optChangeSexMale.Value = True
            fraDialogue.Visible = True
            fraCommand(15).Visible = True
            fraCommands.Visible = False
        Case 20
            optChangePKYes.Value = True
            fraDialogue.Visible = True
            fraCommand(16).Visible = True
            fraCommands.Visible = False
        Case 21
            scrlGiveExp.Value = 0
            lblGiveExp.Caption = "Dar Exp: 0"
            fraDialogue.Visible = True
            fraCommand(17).Visible = True
            fraCommands.Visible = False
        Case 22
            scrlWPMap.Value = 0
            scrlWPX.Value = 0
            scrlWPY.Value = 0
            cmbWarpPlayerDir.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(18).Visible = True
            fraCommands.Visible = False
        Case 23
            fraMoveRoute.Visible = True
            lstMoveRoute.Clear
            cmbEvent.Clear
            lstCommands.Visible = False 'fixeado por EaSee Community
            ReDim ListOfEvents(0 To Map.EventCount)
            ListOfEvents(0) = EditorEvent
            cmbEvent.AddItem "Este Evento"
            cmbEvent.ListIndex = 0
            cmbEvent.Enabled = True
            For I = 1 To Map.EventCount
                If I <> EditorEvent Then
                    cmbEvent.AddItem Trim$(Map.Events(I).name)
                    X = X + 1
                    ListOfEvents(X) = I
                End If
            Next
            IsMoveRouteCommand = True
            chkIgnoreMove.Value = 0
            chkRepeatRoute.Value = 0
            TempMoveRouteCount = 0
            ReDim TempMoveRoute(0)
            fraMoveRoute.Width = 841
            fraMoveRoute.Height = 585
            fraMoveRoute.Visible = True
            fraCommands.Visible = False
        Case 24
            cmbSpawnNPC.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(19).Visible = True
            fraCommands.Visible = False
        Case 25
            cmbPlayAnimEvent.Clear
            For I = 1 To Map.EventCount
                cmbPlayAnimEvent.AddItem I & ". " & Trim$(Map.Events(I).name)
            Next
            cmbPlayAnimEvent.ListIndex = 0
            optPlayAnimPlayer.Value = True
            cmbPlayAnim.ListIndex = 0
            lblPlayAnimX.Caption = "Map Tile X: 0"
            lblPlayAnimY.Caption = "Map Tile Y: 0"
            scrlPlayAnimTileX.Value = 0
            scrlPlayAnimTileY.Value = 0
            scrlPlayAnimTileX.max = Map.MaxX
            scrlPlayAnimTileY.max = Map.MaxY
            fraDialogue.Visible = True
            fraCommand(20).Visible = True
            fraCommands.Visible = False
            lblPlayAnimX.Visible = False
            lblPlayAnimY.Visible = False
            scrlPlayAnimTileX.Visible = False
            scrlPlayAnimTileY.Visible = False
            cmbPlayAnimEvent.Visible = False
        Case 26
            AddCommand EventType.evOpenBank
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 27
            fraDialogue.Visible = True
            fraCommand(21).Visible = True
            cmbOpenShop.ListIndex = 0
            fraCommands.Visible = False
        Case 28
            AddCommand EventType.evFadeIn
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 29
            AddCommand EventType.evFadeOut
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 30
            AddCommand EventType.evFlashWhite
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 31
            ScrlFogData(0).Value = 0
            ScrlFogData(0).Value = 0
            ScrlFogData(0).Value = 0
            fraDialogue.Visible = True
            fraCommand(22).Visible = True
            fraCommands.Visible = False
        Case 32
            CmbWeather.ListIndex = 0
            scrlWeatherIntensity.Value = 0
            fraDialogue.Visible = True
            fraCommand(23).Visible = True
            fraCommands.Visible = False
        Case 33
            scrlMapTintData(0).Value = 0
            scrlMapTintData(1).Value = 0
            scrlMapTintData(2).Value = 0
            scrlMapTintData(3).Value = 0
            fraDialogue.Visible = True
            fraCommand(24).Visible = True
            fraCommands.Visible = False
        Case 34
            cmbPlayBGM.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(25).Visible = True
            fraCommands.Visible = False
        Case 35
            AddCommand EventType.evFadeoutBGM
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 36
            cmbPlaySound.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(26).Visible = True
            fraCommands.Visible = False
        Case 37
            AddCommand EventType.evStopSound
            fraCommands.Visible = False
            fraDialogue.Visible = False
        Case 38
            scrlWaitAmount.Value = 1
            fraDialogue.Visible = True
            fraCommand(27).Visible = True
            fraCommands.Visible = False
        Case 39
            cmbSetAccess.ListIndex = 0
            fraDialogue.Visible = True
            fraCommand(28).Visible = True
            fraCommands.Visible = False
        Case 40
            scrlCustomScript.Value = 1
            fraDialogue.Visible = True
            fraCommand(29).Visible = True
            fraCommands.Visible = False
    End Select
End Sub

Private Sub cmdCondition_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(7).Visible = False
End Sub

Private Sub cmdCondition_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evCondition
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommands.Visible = False
    fraCommand(7).Visible = False
End Sub

Private Sub cmdCopyPage_Click()
    CopyMemory ByVal VarPtr(copyPage), ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    cmdPastePage.Enabled = True
End Sub

Private Sub cmdCreateLabel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(8).Visible = False
End Sub

Private Sub cmdCreateLabel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evLabel
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(8).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdCustomScript_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(29).Visible = False
End Sub

Private Sub cmdCustomScript_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evCustomScript
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(29).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdDeleteCommand_Click()
    If lstCommands.ListIndex > -1 Then
        DeleteEventCommand
    End If
End Sub

Private Sub cmdDeletePage_Click()
Dim I As Long
    ZeroMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), LenB(tmpEvent.Pages(curPageNum))
    ' move everything else down a notch
    If curPageNum < tmpEvent.pageCount Then
        For I = curPageNum To tmpEvent.pageCount - 1
            CopyMemory ByVal VarPtr(tmpEvent.Pages(I)), ByVal VarPtr(tmpEvent.Pages(I + 1)), LenB(tmpEvent.Pages(I + 1))
        Next
    End If
    tmpEvent.pageCount = tmpEvent.pageCount - 1
    ' set the tabs
    tabPages.Tabs.Clear
    For I = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , str(I)
    Next
    ' set the tab back
    If curPageNum <= tmpEvent.pageCount Then
        tabPages.SelectedItem = tabPages.Tabs(curPageNum)
    Else
        tabPages.SelectedItem = tabPages.Tabs(tmpEvent.pageCount)
    End If
    ' make sure we disable
    If tmpEvent.pageCount <= 1 Then
        cmdDeletePage.Enabled = False
    End If
End Sub

Private Sub cmdEditCommand_Click()
Dim I As Long
If lstCommands.ListIndex > -1 Then
    EditEventCommand
End If
End Sub

Private Sub cmdGiveExp_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(17).Visible = False
End Sub

Private Sub cmdGiveExp_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evGiveExp
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(17).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdGotoLabel_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(9).Visible = False
End Sub

Private Sub cmdGotoLabel_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evGotoLabel
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(9).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdGraphicCancel_Click()
    fraGraphic.Visible = False
    lstCommands.Visible = True
End Sub

Private Sub cmdGraphicOK_Click()
    If GraphicSelType = 0 Then
        tmpEvent.Pages(curPageNum).GraphicType = cmbGraphic.ListIndex
        tmpEvent.Pages(curPageNum).Graphic = scrlGraphic.Value
        tmpEvent.Pages(curPageNum).GraphicX = GraphicSelX
        tmpEvent.Pages(curPageNum).GraphicY = GraphicSelY
        tmpEvent.Pages(curPageNum).GraphicX2 = GraphicSelX2
        tmpEvent.Pages(curPageNum).GraphicY2 = GraphicSelY2
    Else
        AddMoveRouteCommand 42
        GraphicSelType = 0
    End If
    fraGraphic.Visible = False
    lstCommands.Visible = True
End Sub

Private Sub cmdLabel_Cancel_Click()
    fraLabeling.Visible = False
    lstCommands.Visible = True
    RequestSwitchesAndVariables
End Sub

Private Sub cmdLabel_Click()
Dim I As Long
    fraLabeling.Visible = True
    lstCommands.Visible = False
    fraLabeling.Width = 849
    fraLabeling.Height = 593
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdMapTint_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(24).Visible = False
End Sub

Private Sub cmdMapTint_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetTint
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(24).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdMoveRoute_Click()
Dim I As Long
    lstCommands.Visible = False
    fraMoveRoute.Visible = True
    lstMoveRoute.Clear
    cmbEvent.Clear
    cmbEvent.AddItem "This Event"
    cmbEvent.ListIndex = 0
    cmbEvent.Enabled = False
    
    IsMoveRouteCommand = False
    
    chkIgnoreMove.Value = tmpEvent.Pages(curPageNum).IgnoreMoveRoute
    chkRepeatRoute.Value = tmpEvent.Pages(curPageNum).RepeatMoveRoute
    
    TempMoveRouteCount = tmpEvent.Pages(curPageNum).MoveRouteCount
    'Will it let me do this?
    TempMoveRoute = tmpEvent.Pages(curPageNum).MoveRoute
    
    For I = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(I).Index
            Case 1
                lstMoveRoute.AddItem "Mover Arriba"
            Case 2
                lstMoveRoute.AddItem "Mover Abajo"
            Case 3
                lstMoveRoute.AddItem "Mover Izquierda"
            Case 4
                lstMoveRoute.AddItem "Mover Derecha"
            Case 5
                lstMoveRoute.AddItem "Mover Azar"
            Case 6
                lstMoveRoute.AddItem "Mover Hacia Jugador"
            Case 7
                lstMoveRoute.AddItem "Alejarse Jugador"
            Case 8
                lstMoveRoute.AddItem "Paso Adelante"
            Case 9
                lstMoveRoute.AddItem "Paso Atras"
            Case 10
                lstMoveRoute.AddItem "Esperar 100ms"
            Case 11
                lstMoveRoute.AddItem "Esperar 500ms"
            Case 12
                lstMoveRoute.AddItem "Esperar 1s"
            Case 13
                lstMoveRoute.AddItem "Mirar Arriba"
            Case 14
                lstMoveRoute.AddItem "Mirar Abajo"
            Case 15
                lstMoveRoute.AddItem "Mirar Izq"
            Case 16
                lstMoveRoute.AddItem "Mirar Der"
            Case 17
                lstMoveRoute.AddItem "Girar 90° a la Derecha"
            Case 18
                lstMoveRoute.AddItem "Girar 90° a la Izquierda"
            Case 19
                lstMoveRoute.AddItem "Girar 180°"
            Case 20
                lstMoveRoute.AddItem "Girar al Azar"
            Case 21
                lstMoveRoute.AddItem "Girar Hacia Jugador"
            Case 22
                lstMoveRoute.AddItem "Girar Contra Jugador"
            Case 23
                lstMoveRoute.AddItem "Ralentizar 8x"
            Case 24
                lstMoveRoute.AddItem "Ralentizar 4x"
            Case 25
                lstMoveRoute.AddItem "Ralentizar 2x"
            Case 26
                lstMoveRoute.AddItem "Velocidad Normal"
            Case 27
                lstMoveRoute.AddItem "Acelerar 2x"
            Case 28
                lstMoveRoute.AddItem "Acelerar 4x"
            Case 29
                lstMoveRoute.AddItem "Frecuencia Menor"
            Case 30
                lstMoveRoute.AddItem "Frecuencia Minima"
            Case 31
                lstMoveRoute.AddItem "Frecuencia Normal"
            Case 32
                lstMoveRoute.AddItem "Frecuencia Mayor"
            Case 33
                lstMoveRoute.AddItem "Frecuencia Maxima"
            Case 34
                lstMoveRoute.AddItem "Animacion al Caminar"
            Case 35
                lstMoveRoute.AddItem "Sin Animacion"
            Case 36
                lstMoveRoute.AddItem "Corregir Direccion"
            Case 37
                lstMoveRoute.AddItem "No Corregir Dir"
            Case 38
                lstMoveRoute.AddItem "Pasar a traves"
            Case 39
                lstMoveRoute.AddItem "No atravezar"
            Case 40
                lstMoveRoute.AddItem "Posicionar debajo de Jugador"
            Case 41
                lstMoveRoute.AddItem "Posicionar en Jugador"
            Case 42
                lstMoveRoute.AddItem "Posicionar arriba de Jugador"
            Case 43
                lstMoveRoute.AddItem "Sprite..."
        End Select
    Next
    
    fraMoveRoute.Width = 841
    fraMoveRoute.Height = 585
    fraMoveRoute.Visible = True
    
End Sub

Sub PopulateMoveRouteList()
Dim I As Long
    lstMoveRoute.Clear
    For I = 1 To TempMoveRouteCount
        Select Case TempMoveRoute(I).Index
            Case 1
                lstMoveRoute.AddItem "Mover Arriba"
            Case 2
                lstMoveRoute.AddItem "Mover Abajo"
            Case 3
                lstMoveRoute.AddItem "Mover Izquierda"
            Case 4
                lstMoveRoute.AddItem "Mover Derecha"
            Case 5
                lstMoveRoute.AddItem "Mover Azar"
            Case 6
                lstMoveRoute.AddItem "Mover Hacia Jugador"
            Case 7
                lstMoveRoute.AddItem "Alejarse Jugador"
            Case 8
                lstMoveRoute.AddItem "Paso Adelante"
            Case 9
                lstMoveRoute.AddItem "Paso Atras"
            Case 10
                lstMoveRoute.AddItem "Esperar 100ms"
            Case 11
                lstMoveRoute.AddItem "Esperar 500ms"
            Case 12
                lstMoveRoute.AddItem "Esperar 1s"
            Case 13
                lstMoveRoute.AddItem "Mirar Arriba"
            Case 14
                lstMoveRoute.AddItem "Mirar Abajo"
            Case 15
                lstMoveRoute.AddItem "Mirar Izq"
            Case 16
                lstMoveRoute.AddItem "Mirar Der"
            Case 17
                lstMoveRoute.AddItem "Girar 90° a la Derecha"
            Case 18
                lstMoveRoute.AddItem "Girar 90° a la Izquierda"
            Case 19
                lstMoveRoute.AddItem "Girar 180°"
            Case 20
                lstMoveRoute.AddItem "Girar al Azar"
            Case 21
                lstMoveRoute.AddItem "Girar Hacia Jugador"
            Case 22
                lstMoveRoute.AddItem "Girar Contra Jugador"
            Case 23
                lstMoveRoute.AddItem "Ralentizar 8x"
            Case 24
                lstMoveRoute.AddItem "Ralentizar 4x"
            Case 25
                lstMoveRoute.AddItem "Ralentizar 2x"
            Case 26
                lstMoveRoute.AddItem "Velocidad Normal"
            Case 27
                lstMoveRoute.AddItem "Acelerar 2x"
            Case 28
                lstMoveRoute.AddItem "Acelerar 4x"
            Case 29
                lstMoveRoute.AddItem "Frecuencia Menor"
            Case 30
                lstMoveRoute.AddItem "Frecuencia Minima"
            Case 31
                lstMoveRoute.AddItem "Frecuencia Normal"
            Case 32
                lstMoveRoute.AddItem "Frecuencia Mayor"
            Case 33
                lstMoveRoute.AddItem "Frecuencia Maxima"
            Case 34
                lstMoveRoute.AddItem "Animacion al Caminar"
            Case 35
                lstMoveRoute.AddItem "Sin Animacion"
            Case 36
                lstMoveRoute.AddItem "Corregir Direccion"
            Case 37
                lstMoveRoute.AddItem "No Corregir Dir"
            Case 38
                lstMoveRoute.AddItem "Pasar a traves"
            Case 39
                lstMoveRoute.AddItem "No atravezar"
            Case 40
                lstMoveRoute.AddItem "Posicionar debajo de Jugador"
            Case 41
                lstMoveRoute.AddItem "Posicionar en Jugador"
            Case 42
                lstMoveRoute.AddItem "Posicionar arriba de Jugador"
            Case 43
                lstMoveRoute.AddItem "Sprite..."
        End Select
    Next
End Sub

Private Sub cmdMoveRouteCancel_Click()
    TempMoveRouteCount = 0
    ReDim TempMoveRoute(0)
    fraMoveRoute.Visible = False
    lstCommands.Visible = True
End Sub

Private Sub cmdMoveRouteOk_Click()
    If IsMoveRouteCommand = True Then
        If Not isEdit Then
            AddCommand EventType.evSetMoveRoute
        Else
            EditCommand
        End If
        TempMoveRouteCount = 0
        ReDim TempMoveRoute(0)
        fraMoveRoute.Visible = False
        lstCommands.Visible = True
    Else
        tmpEvent.Pages(curPageNum).MoveRouteCount = TempMoveRouteCount
        tmpEvent.Pages(curPageNum).MoveRoute = TempMoveRoute
        TempMoveRouteCount = 0
        ReDim TempMoveRoute(0)
        fraMoveRoute.Visible = False
        lstCommands.Visible = True
    End If
End Sub

Private Sub cmdNewPage_Click()
Dim pageCount As Long, I As Long
    pageCount = tmpEvent.pageCount + 1
    ' redim the array
    ReDim Preserve tmpEvent.Pages(pageCount)
    tmpEvent.pageCount = pageCount
    ' set the tabs
    tabPages.Tabs.Clear
    For I = 1 To tmpEvent.pageCount
        tabPages.Tabs.Add , , str(I)
    Next
    cmdDeletePage.Enabled = True
End Sub

Private Sub cmdOk_Click()
    EventEditorOK
End Sub

Private Sub cmdOpenShop_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(21).Visible = False
End Sub

Private Sub cmdOpenShop_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evOpenShop
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(21).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPastePage_Click()
    CopyMemory ByVal VarPtr(tmpEvent.Pages(curPageNum)), ByVal VarPtr(copyPage), LenB(tmpEvent.Pages(curPageNum))
    EventEditorLoadPage curPageNum
End Sub

Private Sub cmdPlayAnim_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(20).Visible = False
End Sub

Private Sub cmdPlayAnim_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayAnimation
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(20).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayBGM_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(25).Visible = False
End Sub

Private Sub cmdPlayBGM_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayBGM
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(25).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdPlayerSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(5).Visible = False
End Sub

Private Sub cmdPlaySound_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(26).Visible = False
End Sub

Private Sub cmdPlaySound_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evPlaySound
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(26).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdRename_Cancel_Click()
Dim I As Long
    fraRenaming.Visible = False
    RenameType = 0
    RenameIndex = 0
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRename_Ok_Click()
Dim I As Long
    Select Case RenameType
        Case 1
            'Variable
            If RenameIndex > 0 And RenameIndex <= MAX_VARIABLES + 1 Then
                Variables(RenameIndex) = txtRename.text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
        Case 2
            'Switch
            If RenameIndex > 0 And RenameIndex <= MAX_SWITCHES + 1 Then
                Switches(RenameIndex) = txtRename.text
                fraRenaming.Visible = False
                RenameType = 0
                RenameIndex = 0
            End If
    End Select
    
    lstSwitches.Clear
    For I = 1 To MAX_SWITCHES
        lstSwitches.AddItem CStr(I) & ". " & Trim$(Switches(I))
    Next
    lstSwitches.ListIndex = 0
    
    lstVariables.Clear
    For I = 1 To MAX_VARIABLES
        lstVariables.AddItem CStr(I) & ". " & Trim$(Variables(I))
    Next
    lstVariables.ListIndex = 0
End Sub

Private Sub cmdRenameSwitch_Click()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editando Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub cmdRenameVariable_Click()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editando Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub cmdSelfSwitch_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(6).Visible = False
End Sub

Private Sub cmdSelfSwitch_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSelfSwitch
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(6).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetAccess_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(28).Visible = False
End Sub

Private Sub cmdSetAccess_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetAccess
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(28).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetFog_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(22).Visible = False
End Sub

Private Sub cmdSetFog_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetFog
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(22).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSetWeather_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(23).Visible = False
End Sub

Private Sub cmdSetWeather_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evSetWeather
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(23).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdShowChatBubble_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(3).Visible = False
End Sub

Private Sub cmdShowChatBubble_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowChatBubble
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(3).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdShowText_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(0).Visible = False
End Sub

Private Sub cmdShowText_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evShowText
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(0).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdSpawnNpc_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(19).Visible = False
End Sub

Private Sub cmdSpawnNpc_Ok_Click()
    If isEdit = False Then
        AddCommand EventType.evSpawnNpc
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(19).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdVariableCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(4).Visible = False
    lstCommands.Visible = True
End Sub

Private Sub cmdVariableOK_Click()
    If Not isEdit Then
        AddCommand EventType.evPlayerVar
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(4).Visible = False
    fraCommands.Visible = False
    lstCommands.Visible = True
End Sub

Private Sub cmdWait_Cancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(27).Visible = False
End Sub

Private Sub cmdWait_Ok_Click()
    If Not isEdit Then
        AddCommand EventType.evWait
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(27).Visible = False
    fraCommands.Visible = False
End Sub

Private Sub cmdWPCancel_Click()
    If Not isEdit Then fraCommands.Visible = True Else fraCommands.Visible = False
    fraDialogue.Visible = False
    fraCommand(18).Visible = False
End Sub

Private Sub cmdWPOkay_Click()
    If Not isEdit Then
        AddCommand EventType.evWarpPlayer
    Else
        EditCommand
    End If
    ' hide
    fraDialogue.Visible = False
    fraCommand(18).Visible = False
    fraCommands.Visible = False
End Sub
Public Sub InitEventEditorForm()
Dim I As Long
    cmbSwitch.Clear
    For I = 1 To MAX_SWITCHES
        cmbSwitch.AddItem I & ". " & Switches(I)
    Next
    
    cmbSwitch.ListIndex = 0
    
    cmbVariable.Clear
    For I = 1 To MAX_VARIABLES
        cmbVariable.AddItem I & ". " & Variables(I)
    Next
    
    cmbSkilling.Clear
    cmbCondition_SkillReq.Clear
    For I = 1 To MAX_SKILLS
        cmbSkilling.AddItem Trim$(Skill(I).name)
        cmbCondition_SkillReq.AddItem Trim$(Skill(I).name)
    Next
    cmbSkilling.ListIndex = 0
    cmbCondition_SkillReq.ListIndex = 0
    
    cmbVariable.ListIndex = 0
    
    cmbChangeItemIndex.Clear
    For I = 1 To MAX_ITEMS
        cmbChangeItemIndex.AddItem Trim$(Item(I).name)
    Next
    
    cmbChangeItemIndex.ListIndex = 0
    
    scrlChangeLevel.min = 1
    scrlChangeLevel.max = MAX_LEVELS
    scrlChangeLevel.Value = 1
    lblChangeLevel.Caption = "Nivel: 1"
    
    cmbChangeSkills.Clear
    For I = 1 To MAX_SPELLS
        cmbChangeSkills.AddItem Trim$(Spell(I).name)
    Next
    
    cmbChangeSkills.ListIndex = 0
    
    cmbChangeClass.Clear
    If Max_Classes > 0 Then
        For I = 1 To Max_Classes
            cmbChangeClass.AddItem Trim$(Class(I).name)
        Next
        cmbChangeClass.ListIndex = 0
    End If
    
    scrlChangeSprite.max = NumCharacters
    
    cmbPlayAnim.Clear
    For I = 1 To MAX_ANIMATIONS
        cmbPlayAnim.AddItem I & ". " & Trim$(Animation(I).name)
    Next
    cmbPlayAnim.ListIndex = 0
    PopulateLists
    cmbPlayBGM.Clear
    For I = 1 To UBound(musicCache)
        cmbPlayBGM.AddItem (musicCache(I))
    Next
    cmbPlayBGM.ListIndex = 0
    
    cmbPlaySound.Clear
    For I = 1 To UBound(soundCache)
        cmbPlaySound.AddItem (soundCache(I))
    Next
    cmbPlaySound.ListIndex = 0
    
    cmbOpenShop.Clear
    For I = 1 To MAX_SHOPS
        cmbOpenShop.AddItem I & ". " & Trim$(Shop(I).name)
    Next
    
    cmbOpenShop.ListIndex = 0
    
    
    cmbSpawnNPC.Clear
    For I = 1 To MAX_MAP_NPCS
        If Map.NPC(I) > 0 Then
            cmbSpawnNPC.AddItem I & ". " & Trim$(NPC(Map.NPC(I)).name)
        Else
            cmbSpawnNPC.AddItem I & ". "
        End If
    Next
    
    cmbSpawnNPC.ListIndex = 0
    
    ScrlFogData(0).max = NumFogs
End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "EventEditor", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L1")
fraRandom(20).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L2")
lblRandomLabel(32).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L3")
fraRandom(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L4")
chkPlayerVar.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L5")
lblRandomLabel(5).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L6")
chkPlayerSwitch.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L7")
chkHasItem.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L8")
chkSelfSwitch.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L9")
fraRandom(13).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L10")
fraRandom(15).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L11")
lblRandomLabel(6).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L12")
lblRandomLabel(7).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L13")
lblRandomLabel(8).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L14")
fraRandom(19).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L15")
fraRandom(16).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L16")
chkWalkAnim.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L17")
chkDirFix.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L18")
chkWalkThrough.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L19")
chkShowName.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L20")
fraRandom(18).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L21")
fraRandom(17).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L22")
chkGlobal.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L23")
fraCommand(7).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L25")
optCondition_Index(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L26")
lblRandomLabel(0).Caption = lblRandomLabel(5).Caption
optCondition_Index(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L27")
optCondition_Index(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L28")
optCondition_Index(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L29")
optCondition_Index(4).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L30")
optCondition_Index(5).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L31")
optCondition_Index(6).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L32")
optCondition_Index(7).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L33")
optCondition_Index(8).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L34")
lblRandomLabel(46).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L35")
fraCommand(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L36")
lblAddText_Colour.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L37")
lblRandomLabel(10).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L38")
optAddText_Player.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L39")
optAddText_Map.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L40")
optAddText_Global.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L41")
fraCommand(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L42")
lblRandomLabel(34).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L43")
lblRandomLabel(39).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L44")
optChatBubbleTarget(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L45")
optChatBubbleTarget(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L46")
optChatBubbleTarget(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L47")
fraCommand(4).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L48")
lblRandomLabel(12).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L49")
optVariableAction(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L50")
optVariableAction(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L51")
optVariableAction(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L52")
optVariableAction(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L53")
lblRandomLabel(13).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L54")
lblRandomLabel(37).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L55")
fraCommand(5).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L56")
lblRandomLabel(22).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L57")
fraCommand(6).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L58")
lblRandomLabel(24).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L59")
lblRandomLabel(26).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L60")
fraCommand(8).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L61")
lblRandomLabel(40).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L62")
fraCommand(9).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L63")
lblRandomLabel(41).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L64")
fraCommand(10).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L65")
lblRandomLabel(27).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L66")
optChangeItemSet.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L67")
optChangeItemAdd.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L68")
optChangeItemRemove.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L69")
fraCommand(11).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L70")
lblChangeLevel.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L71")
fraCommand(12).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L72")
lblRandomLabel(28).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L73")
optChangeSkillsAdd.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L74")
optChangeSkillsRemove.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L75")
fraCommand(13).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L76")
lblRandomLabel(29).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L77")
fraCommand(14).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L78")
fraCommand(15).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L79")
optChangeSexMale.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L80")
optChangeSexFemale.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L81")
optChangePKYes.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L82")
optChangePKNo.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L83")
fraCommand(17).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L84")
opMine.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L85")
opSkilling.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L86")
lblGiveExp.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L87")
fraCommand(18).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L88")
lblWPMap.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L89")
fraCommand(20).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L90")
lblRandomLabel(30).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L91")
lblRandomLabel(31).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L92")
optPlayAnimPlayer.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L93")
optPlayAnimEvent.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L94")
fraCommand(21).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L95")
fraCommand(22).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L96")
lblFogData(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L97")
lblFogData(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L98")
lblFogData(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L99")
fraCommand(23).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L100")
lblRandomLabel(43).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L101")
lblWeatherIntensity.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L102")
fraCommand(24).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L103")
lblMapTintData(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L104")
lblMapTintData(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L105")
lblMapTintData(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L106")
lblMapTintData(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L107")
fraCommand(26).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L108")
fraCommand(27).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L109")
lblWaitAmount.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L110")
lblRandomLabel(44).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L111")
fraCommand(28).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L112")
lblCustomScript.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L113")
fraCommand(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L114")
lblRandomLabel(18).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L115")
fraCommand(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L116")
lblRandomLabel(16).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L117")
lblRandomLabel(17).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L118")
lblRandomLabel(19).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L119")
lblRandomLabel(20).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L120")
lblRandomLabel(21).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L121")
fraCommands.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L122")
fraRandom(21).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L123")
fraRandom(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L124")
fraRandom(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "L125")
fraRandom(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B1")
cmdNewPage.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B2")
cmdCopyPage.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B3")
cmdPastePage.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B4")
cmdDeletePage.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B5")
cmdClearPage.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B6")
cmdLabel.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B7")
cmdAddCommand.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B8")
cmdEditCommand.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B9")
cmdDeleteCommand.Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "B10")
cmdClearCommand.Caption = trad

trad = GetVar(App.Path & Lang, "EventEditor", "CB1")
cmdCommands(0).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB2")
cmdCommands(1).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB3")
cmdCommands(2).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB4")
cmdCommands(3).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB5")
cmdCommands(4).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB6")
cmdCommands(5).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB7")
cmdCommands(6).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB8")
cmdCommands(7).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB9")
cmdCommands(8).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB10")
cmdCommands(9).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB11")
cmdCommands(10).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB12")
cmdCommands(11).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB13")
cmdCommands(12).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB14")
cmdCommands(13).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB15")
cmdCommands(14).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB16")
cmdCommands(15).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB17")
cmdCommands(16).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB18")
cmdCommands(17).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB19")
cmdCommands(18).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB20")
cmdCommands(19).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB21")
cmdCommands(20).Caption = trad
trad = GetVar(App.Path & Lang, "EventEditor", "CB22")
cmdCommands(21).Caption = trad

    InitEventEditorForm
End Sub

Private Sub lstCommands_Click()
    curCommand = lstCommands.ListIndex + 1
End Sub

Sub AddMoveRouteCommand(Index As Integer)
Dim I As Long, X As Long, Z As Long
    Index = Index + 1
    If lstMoveRoute.ListIndex > -1 Then
        I = lstMoveRoute.ListIndex + 1
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        For X = TempMoveRouteCount - 1 To I Step -1
            TempMoveRoute(X + 1) = TempMoveRoute(X)
        Next
        TempMoveRoute(I).Index = Index
        'if set graphic then...
        If Index = 43 Then
            TempMoveRoute(I).Data1 = cmbGraphic.ListIndex
            TempMoveRoute(I).Data2 = scrlGraphic.Value
            TempMoveRoute(I).Data3 = GraphicSelX
            TempMoveRoute(I).Data4 = GraphicSelX2
            TempMoveRoute(I).Data5 = GraphicSelY
            TempMoveRoute(I).Data6 = GraphicSelY2
        End If
        PopulateMoveRouteList
    Else
        TempMoveRouteCount = TempMoveRouteCount + 1
        ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        TempMoveRoute(TempMoveRouteCount).Index = Index
        PopulateMoveRouteList
        'if set graphic then....
        If Index = 43 Then
            TempMoveRoute(TempMoveRouteCount).Data1 = cmbGraphic.ListIndex
            TempMoveRoute(TempMoveRouteCount).Data2 = scrlGraphic.Value
            TempMoveRoute(TempMoveRouteCount).Data3 = GraphicSelX
            TempMoveRoute(TempMoveRouteCount).Data4 = GraphicSelX2
            TempMoveRoute(TempMoveRouteCount).Data5 = GraphicSelY
            TempMoveRoute(TempMoveRouteCount).Data6 = GraphicSelY2
        End If
    End If
End Sub

Sub RemoveMoveRouteCommand(Index As Long)
Dim I As Long
    Index = Index + 1
    If Index > 0 And Index <= TempMoveRouteCount Then
        For I = Index + 1 To TempMoveRouteCount
            TempMoveRoute(I - 1) = TempMoveRoute(I)
        Next
        TempMoveRouteCount = TempMoveRouteCount - 1
        If TempMoveRouteCount = 0 Then
            ReDim TempMoveRoute(0)
        Else
            ReDim Preserve TempMoveRoute(TempMoveRouteCount)
        End If
        PopulateMoveRouteList
    End If
End Sub

Private Sub lstCommands_DblClick()
    cmdAddCommand_Click
End Sub

Private Sub lstCommands_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        'remove move route command lol
        cmdDeleteCommand_Click
    End If
End Sub

Private Sub lstMoveRoute_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        'remove move route command lol
        If lstMoveRoute.ListIndex > -1 Then
            Call RemoveMoveRouteCommand(lstMoveRoute.ListIndex)
        End If
    End If
End Sub

Private Sub optAddText_Game_Click()

End Sub

Private Sub lstSwitches_DblClick()
    If lstSwitches.ListIndex > -1 And lstSwitches.ListIndex < MAX_SWITCHES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editando Switch #" & CStr(lstSwitches.ListIndex + 1)
        txtRename.text = Switches(lstSwitches.ListIndex + 1)
        RenameType = 2
        RenameIndex = lstSwitches.ListIndex + 1
    End If
End Sub

Private Sub lstVariables_DblClick()
    If lstVariables.ListIndex > -1 And lstVariables.ListIndex < MAX_VARIABLES Then
        fraRenaming.Visible = True
        lblEditing.Caption = "Editando Variable #" & CStr(lstVariables.ListIndex + 1)
        txtRename.text = Variables(lstVariables.ListIndex + 1)
        RenameType = 1
        RenameIndex = lstVariables.ListIndex + 1
    End If
End Sub

Private Sub opMine_Click()
    cmbSkilling.Visible = False
End Sub

Private Sub opSkilling_Click()
    cmbSkilling.Visible = True
    cmbSkilling.ListIndex = 0
End Sub

Private Sub optChatBubbleTarget_Click(Index As Integer)
Dim I As Long
    If Index = 0 Then
        cmbChatBubbleTarget.Visible = False
    ElseIf Index = 1 Then
        cmbChatBubbleTarget.Visible = True
        cmbChatBubbleTarget.Clear
        For I = 1 To MAX_MAP_NPCS
            If Map.NPC(I) <= 0 Then
                cmbChatBubbleTarget.AddItem CStr(I) & ". "
            Else
                cmbChatBubbleTarget.AddItem CStr(I) & ". " & Trim$(NPC(Map.NPC(I)).name)
            End If
        Next
        cmbChatBubbleTarget.ListIndex = 0
    ElseIf Index = 2 Then
        cmbChatBubbleTarget.Visible = True
        cmbChatBubbleTarget.Clear
        For I = 1 To Map.EventCount
            cmbChatBubbleTarget.AddItem CStr(I) & ". " & Trim$(Map.Events(I).name)
        Next
        cmbChatBubbleTarget.ListIndex = 0
    End If
End Sub

Private Sub optCondition_Index_Click(Index As Integer)
Dim I As Long, X As Long
    For I = 0 To 8
        If optCondition_Index(I).Value = True Then X = I
    Next
    ClearConditionFrame
    Select Case X
        Case 0
            cmbCondition_PlayerVarIndex.Enabled = True
            cmbCondition_PlayerVarCompare.Enabled = True
            txtCondition_PlayerVarCondition.Enabled = True
        Case 1
            cmbCondition_PlayerSwitch.Enabled = True
            cmbCondtion_PlayerSwitchCondition.Enabled = True
        Case 2
            cmbCondition_HasItem.Enabled = True
            txtCondition_itemAmount.Enabled = True
        Case 3
            cmbCondition_ClassIs.Enabled = True
        Case 4
            cmbCondition_LearntSkill.Enabled = True
        Case 5
            cmbCondition_LevelCompare.Enabled = True
            txtCondition_LevelAmount.Enabled = True
        Case 6
            cmbCondition_SelfSwitch.Enabled = True
            cmbCondition_SelfSwitchCondition.Enabled = True
        Case 7
            cmbCondition_SkillReq.Enabled = True
            txtCondition_SkillLvlReq.Enabled = True
        Case 8
            cmbCondition_Quest.Enabled = True
            cmbCondition_Status.Enabled = True
    End Select
End Sub
Sub ClearConditionFrame()
Dim I As Long, tempCheck As String * 30
    cmbCondition_PlayerVarIndex.Enabled = False
    cmbCondition_PlayerVarIndex.Clear
    For I = 1 To MAX_VARIABLES
        cmbCondition_PlayerVarIndex.AddItem I & ". " & Variables(I)
    Next
    cmbCondition_PlayerVarIndex.ListIndex = 0
    
    cmbCondition_PlayerVarCompare.ListIndex = 0
    cmbCondition_PlayerVarCompare.Enabled = False
    
    cmbCondition_Status.ListIndex = 0
    cmbCondition_Status.Enabled = False
    cmbCondition_Quest.Clear
    For I = 1 To MAX_QUESTS
        cmbCondition_Quest.AddItem CStr(I) & ": " & Trim$(Quest(I).name)
    Next I
    
    cmbCondition_Quest.ListIndex = 0
    cmbCondition_Quest.Enabled = False
    
    txtCondition_PlayerVarCondition.Enabled = False
    txtCondition_PlayerVarCondition.text = "0"
    
    cmbCondition_PlayerSwitch.Enabled = False
    cmbCondition_PlayerSwitch.Clear
    For I = 1 To MAX_SWITCHES
        cmbCondition_PlayerSwitch.AddItem I & ". " & Switches(I)
    Next
    cmbCondition_PlayerSwitch.ListIndex = 0
    
    cmbCondtion_PlayerSwitchCondition.Enabled = False
    cmbCondtion_PlayerSwitchCondition.ListIndex = 0
    
    cmbCondition_HasItem.Enabled = False
    cmbCondition_HasItem.Clear
    For I = 1 To MAX_ITEMS
        cmbCondition_HasItem.AddItem I & ". " & Trim$(Item(I).name)
    Next
    cmbCondition_HasItem.ListIndex = 0
    
    txtCondition_itemAmount.Enabled = False
    txtCondition_itemAmount.text = "0"
    
    cmbCondition_ClassIs.Enabled = False
    cmbCondition_ClassIs.Clear
    For I = 1 To Max_Classes
        cmbCondition_ClassIs.AddItem I & ". " & CStr(Class(I).name)
    Next
    cmbCondition_ClassIs.ListIndex = 0
    
    cmbCondition_LearntSkill.Enabled = False
    cmbCondition_LearntSkill.Clear
    For I = 1 To MAX_SPELLS
        cmbCondition_LearntSkill.AddItem I & ". " & Trim$(Spell(I).name)
    Next
    
    cmbCondition_SkillReq.Clear
    For I = 1 To MAX_SKILLS
        cmbCondition_SkillReq.AddItem Trim$(Skill(I).name)
    Next
    cmbCondition_SkillReq.ListIndex = 0
    cmbCondition_SkillReq.Enabled = False
    txtCondition_SkillLvlReq.Enabled = False
    cmbCondition_LearntSkill.ListIndex = 0
    cmbCondition_LevelCompare.Enabled = False
    cmbCondition_LevelCompare.ListIndex = 0
    txtCondition_LevelAmount.Enabled = False
    txtCondition_LevelAmount.text = "0"
    
    cmbCondition_SelfSwitch.ListIndex = 0
    cmbCondition_SelfSwitch.Enabled = False
    cmbCondition_SelfSwitchCondition.ListIndex = 0
    cmbCondition_SelfSwitchCondition.Enabled = False
End Sub

Private Sub optPlayAnimEvent_Click()
    lblPlayAnimX.Visible = False
    lblPlayAnimY.Visible = False
    scrlPlayAnimTileX.Visible = False
    scrlPlayAnimTileY.Visible = False
    cmbPlayAnimEvent.Visible = True
End Sub

Private Sub optPlayAnimPlayer_Click()
    lblPlayAnimX.Visible = False
    lblPlayAnimY.Visible = False
    scrlPlayAnimTileX.Visible = False
    scrlPlayAnimTileY.Visible = False
    cmbPlayAnimEvent.Visible = False
End Sub

Private Sub optPlayAnimTile_Click()
    lblPlayAnimX.Visible = True
    lblPlayAnimY.Visible = True
    scrlPlayAnimTileX.Visible = True
    scrlPlayAnimTileY.Visible = True
    cmbPlayAnimEvent.Visible = False
End Sub

Private Sub optVariableAction_Click(Index As Integer)
    Dim I As Long
    For I = 0 To 3
        If optVariableAction(I).Value = True Then
            Exit For
        End If
    Next
    txtVariableData(0).Enabled = False
    txtVariableData(0).text = "0"
    txtVariableData(1).Enabled = False
    txtVariableData(1).text = "0"
    txtVariableData(2).Enabled = False
    txtVariableData(2).text = "0"
    txtVariableData(3).Enabled = False
    txtVariableData(3).text = "0"
    txtVariableData(4).Enabled = False
    txtVariableData(4).text = "0"
    Select Case I
        Case 0
            txtVariableData(0).Enabled = True
        Case 1
            txtVariableData(1).Enabled = True
        Case 2
            txtVariableData(2).Enabled = True
        Case 3
            txtVariableData(3).Enabled = True
            txtVariableData(4).Enabled = True
    End Select
End Sub

Private Sub picGraphic_Click()
    fraGraphic.Width = 841
    fraGraphic.Height = 585
    fraGraphic.Visible = True
    lstCommands.Visible = False
    GraphicSelType = 0
End Sub

Private Sub picGraphicSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim I As Long
    If frmEditor_Events.cmbGraphic.ListIndex = 2 Then
        'Tileset... hard one....
        If ShiftDown Then
            If GraphicSelX > -1 And GraphicSelY > -1 Then
                If CLng(X + frmEditor_Events.hScrlGraphicSel.Value) / 32 > GraphicSelX And CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) / 32 > GraphicSelY Then
                    GraphicSelX2 = CLng(X + frmEditor_Events.hScrlGraphicSel.Value) / 32
                    GraphicSelY2 = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) / 32
                End If
            End If
        Else
            GraphicSelX = CLng(X + frmEditor_Events.hScrlGraphicSel.Value) \ 32
            GraphicSelY = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value) \ 32
            GraphicSelX2 = 0
            GraphicSelY2 = 0
        End If
    ElseIf frmEditor_Events.cmbGraphic.ListIndex = 1 Then
        GraphicSelX = CLng(X + frmEditor_Events.hScrlGraphicSel.Value)
        GraphicSelY = CLng(Y + frmEditor_Events.vScrlGraphicSel.Value)
        GraphicSelX2 = 0
        GraphicSelY2 = 0
        
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
        
        
        If VXFRAME = False Then
            For I = 0 To 3
                If GraphicSelX >= ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) * I) And GraphicSelX < ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) * (I + 1)) Then
                    GraphicSelX = I
                End If
            Next
        Else
            For I = 0 To 2
                If GraphicSelX >= ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 3) * I) And GraphicSelX < ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 3) * (I + 1)) Then
                    GraphicSelX = I
                End If
            Next
        End If
        
        For I = 0 To 3
            If GraphicSelY >= ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) * I) And GraphicSelY < ((Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) * (I + 1)) Then
                GraphicSelY = I
            End If
        Next
        
        
    End If
End Sub

Private Sub scrlGraphic_Click()
    lblGraphic.Caption = "Imagen: " & scrlGraphic.Value
    tmpEvent.Pages(curPageNum).Graphic = scrlGraphic.Value
    
    If tmpEvent.Pages(curPageNum).GraphicType = 1 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumCharacters Then Exit Sub
   
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    ElseIf tmpEvent.Pages(curPageNum).GraphicType = 2 Then
        If frmEditor_Events.scrlGraphic.Value <= 0 Or frmEditor_Events.scrlGraphic.Value > NumTileSets Then Exit Sub
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
            frmEditor_Events.hScrlGraphicSel.Visible = True
            frmEditor_Events.hScrlGraphicSel.Value = 0
            frmEditor_Events.hScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width - 800
        Else
            frmEditor_Events.hScrlGraphicSel.Visible = False
        End If
                    
        If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
            frmEditor_Events.vScrlGraphicSel.Visible = True
            frmEditor_Events.vScrlGraphicSel.Value = 0
            frmEditor_Events.vScrlGraphicSel.max = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height - 512
        Else
            frmEditor_Events.vScrlGraphicSel.Visible = False
        End If
    End If
End Sub

Private Sub scrlAddText_Colour_Click()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
End Sub

Private Sub scrlAddText_Colour_Change()
    frmEditor_Events.lblAddText_Colour.Caption = "Color: " & GetColorString(frmEditor_Events.scrlAddText_Colour.Value)
End Sub

Private Sub scrlChangeLevel_Change()
    lblChangeLevel.Caption = "Nivel: " & scrlChangeLevel.Value
End Sub

Private Sub scrlChangeSprite_Change()
    lblChangeSprite.Caption = "Sprite: " & scrlChangeSprite.Value
End Sub

Private Sub scrlCustomScript_Change()
    lblCustomScript.Caption = "Case: " & scrlCustomScript.Value
End Sub

Private Sub scrlGiveExp_Click()
    lblGiveExp.Caption = "Dar Exp: " & scrlGiveExp.Value
End Sub

Private Sub ScrlFogData_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            If ScrlFogData(0).Value = 0 Then
                lblFogData(0).Caption = "Ninguno."
            Else
                lblFogData(0).Caption = "Niebla: " & ScrlFogData(0).Value
            End If
        Case 1
            lblFogData(1).Caption = "Velocidad Niebla: " & ScrlFogData(1).Value
        Case 2
            lblFogData(2).Caption = "Opacidad Niebla: " & ScrlFogData(2).Value
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScrlFogData_Change(" & CStr(Index) & ")", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGiveExp_Change()
    lblGiveExp.Caption = "Dar Exp: " & scrlGiveExp.Value
End Sub

Private Sub scrlGraphic_Change()
    If scrlGraphic.Value = 0 Then
        lblGraphic.Caption = "Numero: Ninguno"
    Else
        lblGraphic.Caption = "Numero: " & scrlGraphic.Value
    End If
    Call cmbGraphic_Click
End Sub

Private Sub scrlMapTintData_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            lblMapTintData(0).Caption = "Rojo: " & scrlMapTintData(0).Value
        Case 1
            lblMapTintData(1).Caption = "Verde: " & scrlMapTintData(1).Value
        Case 2
            lblMapTintData(2).Caption = "Azul: " & scrlMapTintData(2).Value
        Case 3
            lblMapTintData(3).Caption = "Opacidad: " & scrlMapTintData(3).Value
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScrlMapTintData_Change(" & CStr(Index) & ")", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPlayAnimTileX_Change()
    lblPlayAnimX.Caption = "Tile X: " & scrlPlayAnimTileX.Value
End Sub

Private Sub scrlPlayAnimTileY_Change()
    lblPlayAnimY.Caption = "Tile Y: " & scrlPlayAnimTileY.Value
End Sub

Private Sub scrlWaitAmount_Change()
    lblWaitAmount.Caption = "Espera: " & scrlWaitAmount.Value & " Ms"
End Sub

Private Sub scrlWeatherIntensity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lblWeatherIntensity.Caption = "Intensidad: " & scrlWeatherIntensity.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "ScrlWeatherIntensity_Change", "frmEditor_Events", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlWPMap_Change()
    lblWPMap.Caption = "Mapa: " & scrlWPMap.Value
End Sub

Private Sub scrlWPX_Change()
    lblWPX.Caption = "X: " & scrlWPX.Value
End Sub

Private Sub scrlWPY_Change()
    lblWPY.Caption = "Y: " & scrlWPY.Value
End Sub

Private Sub tabCommands_Click()
Dim I As Long
    For I = 1 To 2
        picCommands(I).Visible = False
    Next
    picCommands(tabCommands.SelectedItem.Index).Visible = True
End Sub

Private Sub tabPages_Click()
    If tabPages.SelectedItem.Index <> curPageNum Then
        curPageNum = tabPages.SelectedItem.Index
        EventEditorLoadPage curPageNum
    End If
End Sub


Private Sub txtName_Validate(Cancel As Boolean)
    tmpEvent.name = Trim$(txtName.text)
End Sub

Private Sub txtPlayerVariable_Validate(Cancel As Boolean)
    tmpEvent.Pages(curPageNum).VariableCondition = Val(Trim$(txtPlayerVariable.text))
End Sub


