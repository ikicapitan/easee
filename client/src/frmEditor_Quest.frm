VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Misiones"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9615
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
   Icon            =   "frmEditor_Quest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   553
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   6720
      TabIndex        =   93
      Top             =   7800
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Título de la Misión"
      Height          =   975
      Left            =   3600
      TabIndex        =   6
      Top             =   120
      Width           =   5895
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recompensas"
         Height          =   180
         Index           =   2
         Left            =   3480
         TabIndex        =   11
         ToolTipText     =   "Recompensas, Premios y Beneficios al Cumplir la Mision o Quest."
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tareas"
         Height          =   180
         Index           =   3
         Left            =   4920
         TabIndex        =   10
         ToolTipText     =   "Tareas a Efectuar para el Cumplimiento de la Quest o Mision."
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Requerimientos"
         Height          =   180
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         ToolTipText     =   "Requerimientos o Condiciones Previas de la Mision a Editar."
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton optShowFrame 
         BackColor       =   &H00E0E0E0&
         Caption         =   "General"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Opciones Generales de la Mision a Editar."
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Ingresa un Titulo para identificar la Mision o Quest."
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "Modificar Longitud"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   2895
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7800
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista Misiones"
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstIndex 
         Height          =   7080
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame fraGeneral 
      BackColor       =   &H00E0E0E0&
      Caption         =   "General"
      Height          =   6495
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdTakeItemRemove 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   4560
         TabIndex        =   74
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdTakeItem 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   3000
         TabIndex        =   73
         Top             =   6120
         Width           =   1575
      End
      Begin VB.ListBox lstTakeItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":08CA
         Left            =   3000
         List            =   "frmEditor_Quest.frx":08CC
         TabIndex        =   71
         Top             =   4080
         Width           =   2775
      End
      Begin VB.ListBox lstGiveItem 
         Height          =   2040
         ItemData        =   "frmEditor_Quest.frx":08CE
         Left            =   120
         List            =   "frmEditor_Quest.frx":08D0
         TabIndex        =   69
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txtQuestLog 
         Height          =   270
         Left            =   1680
         MaxLength       =   200
         TabIndex        =   67
         Top             =   240
         Width           =   4095
      End
      Begin VB.CheckBox chkRepeat 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "Repetitiva"
         Height          =   255
         Left            =   3960
         TabIndex        =   64
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlTakeItem 
         Height          =   255
         Left            =   3000
         Max             =   255
         TabIndex        =   63
         Top             =   3480
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlTakeItemValue 
         Height          =   135
         Left            =   3000
         Max             =   10
         Min             =   1
         TabIndex        =   62
         Top             =   3840
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlGiveItemValue 
         Height          =   135
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   61
         Top             =   3840
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlGiveItem 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   60
         Top             =   3480
         Value           =   1
         Width           =   2775
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   1
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   2
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1800
         Width           =   5655
      End
      Begin VB.TextBox txtSpeech 
         Height          =   270
         Index           =   3
         Left            =   120
         MaxLength       =   250
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2400
         Width           =   5655
      End
      Begin VB.CommandButton cmdGiveItemRemove 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   1680
         TabIndex        =   72
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton cmdGiveItem 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio de Mision:"
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   250
         Width           =   1485
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblTakeItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Quitar Objeto al Finalizar: 0 (1)"
         Height          =   420
         Left            =   3000
         TabIndex        =   66
         Top             =   3000
         Width           =   2745
      End
      Begin VB.Label lblGiveItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Dar Objeto al Iniciar: 0 (1)"
         Height          =   420
         Left            =   120
         TabIndex        =   65
         Top             =   3000
         Width           =   2715
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   120
         X2              =   5160
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label lblQ1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dialogo Peticion:"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lblQ2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Durante Dialogo:"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblQ3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dialogo Finalizar:"
         Height          =   180
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1290
      End
   End
   Begin VB.Frame fraRewards 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recompensas"
      Height          =   6495
      Left            =   3600
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.ComboBox cmbSkill 
         Height          =   300
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlSkillExp 
         Height          =   255
         LargeChange     =   50
         Left            =   3000
         TabIndex        =   94
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton cmdItemRewRemove 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   1680
         TabIndex        =   75
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ListBox lstItemRew 
         Height          =   2220
         ItemData        =   "frmEditor_Quest.frx":08D2
         Left            =   120
         List            =   "frmEditor_Quest.frx":08D9
         TabIndex        =   59
         Top             =   1200
         Width           =   2775
      End
      Begin VB.HScrollBar scrlExp 
         Height          =   255
         LargeChange     =   50
         Left            =   3000
         TabIndex        =   57
         Top             =   600
         Width           =   2775
      End
      Begin VB.HScrollBar scrlItemRew 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   27
         Top             =   600
         Value           =   1
         Width           =   2775
      End
      Begin VB.HScrollBar scrlItemRewValue 
         Height          =   135
         Left            =   120
         Max             =   10
         Min             =   1
         TabIndex        =   26
         Top             =   960
         Value           =   1
         Width           =   2775
      End
      Begin VB.CommandButton cmdItemRew 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Habilidad:"
         Height          =   180
         Left            =   3000
         TabIndex        =   96
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Exp de Habilidad Recompensada: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   95
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recompensa Exp: 0"
         Height          =   180
         Left            =   3000
         TabIndex        =   58
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblItemRew 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto Recompensa: 0 (1)"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1425
      End
   End
   Begin VB.Frame fraTasks 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tareas"
      Height          =   6495
      Left            =   3600
      TabIndex        =   24
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   120
         TabIndex        =   41
         Top             =   600
         Width           =   3375
         Begin VB.HScrollBar scrlEvent 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   91
            Top             =   4080
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.HScrollBar scrlNPC 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   49
            Top             =   1680
            Width           =   3135
         End
         Begin VB.HScrollBar scrlItem 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   48
            Top             =   2280
            Width           =   3135
         End
         Begin VB.HScrollBar scrlAmount 
            Height          =   255
            Left            =   120
            Max             =   10
            TabIndex        =   47
            Top             =   5040
            Width           =   3135
         End
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   46
            Top             =   2880
            Width           =   3135
         End
         Begin VB.TextBox txtTaskSpeech 
            Height          =   270
            Left            =   120
            MaxLength       =   250
            ScrollBars      =   2  'Vertical
            TabIndex        =   45
            ToolTipText     =   "Dialogo para Obtener la Mision o Tarea."
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox txtTaskLog 
            Height          =   270
            Left            =   120
            MaxLength       =   200
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            ToolTipText     =   "Descripcion de la Tarea a Realizar."
            Top             =   1080
            Width           =   3135
         End
         Begin VB.HScrollBar scrlResource 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   43
            Top             =   3480
            Width           =   3135
         End
         Begin VB.CheckBox chkEnd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Finalizar Mision Ahora"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   180
            Left            =   120
            TabIndex        =   42
            Top             =   5400
            Width           =   1935
         End
         Begin VB.Label lblEvent 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Evento: Ninguno"
            Height          =   180
            Left            =   120
            TabIndex        =   92
            Top             =   3840
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lblNPC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NPC: Ninguno"
            Height          =   180
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   1065
         End
         Begin VB.Label lblItem 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto: Ninguno"
            Height          =   180
            Left            =   120
            TabIndex        =   55
            Top             =   2040
            Width           =   1230
         End
         Begin VB.Label lblAmount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   54
            Top             =   4800
            Width           =   885
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   3240
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label lblMap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa: Ninguno"
            Height          =   180
            Left            =   120
            TabIndex        =   53
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label lblSpeech 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Diálogo Tarea:"
            Height          =   180
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label lblLog 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción Tarea:"
            Height          =   180
            Left            =   120
            TabIndex        =   51
            Top             =   840
            Width           =   1440
         End
         Begin VB.Label lblResource 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Recurso Req: Ninguno"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   3240
            Width           =   1665
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   3600
         TabIndex        =   31
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Obtener de Evento"
            Height          =   180
            Index           =   9
            Left            =   120
            TabIndex        =   90
            Top             =   2520
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ninguno"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Derrotar NPC"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Traer Objetos"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Hablar a NPC"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Llegar a Mapa"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dar Objeto a NPC"
            Height          =   180
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Asesinar Jugador"
            Height          =   180
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1695
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Entrenar con Recurso"
            Height          =   180
            Index           =   7
            Left            =   120
            TabIndex        =   33
            Top             =   2040
            Width           =   1935
         End
         Begin VB.OptionButton optTask 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Obtener de NPC"
            Height          =   180
            Index           =   8
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000000&
            X1              =   120
            X2              =   2040
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.HScrollBar scrlTotalTasks 
         Height          =   255
         Left            =   1680
         Max             =   10
         Min             =   1
         TabIndex        =   29
         Top             =   240
         Value           =   1
         Width           =   4095
      End
      Begin VB.Label lblSelected 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarea Selecc: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   360
         TabIndex        =   30
         ToolTipText     =   "Tareas a Efectuar para el Cumplimiento de la Quest o Mision."
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame fraRequirements 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requerimientos"
      Height          =   6495
      Left            =   3600
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   5895
      Begin VB.HScrollBar scrlReqSwitch 
         Height          =   255
         Left            =   120
         Max             =   70
         TabIndex        =   88
         Top             =   1680
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqClass 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   87
         Top             =   2400
         Value           =   1
         Width           =   2415
      End
      Begin VB.ListBox lstReqClass 
         Height          =   1140
         ItemData        =   "frmEditor_Quest.frx":08E9
         Left            =   120
         List            =   "frmEditor_Quest.frx":08EB
         TabIndex        =   86
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqClassRemove 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   1320
         TabIndex        =   84
         Top             =   3960
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqItemValue 
         Height          =   135
         Left            =   2760
         Max             =   10
         Min             =   1
         TabIndex        =   81
         Top             =   840
         Value           =   1
         Width           =   3015
      End
      Begin VB.HScrollBar scrlReqItem 
         Height          =   255
         Left            =   2760
         Max             =   255
         TabIndex        =   80
         Top             =   480
         Value           =   1
         Width           =   3015
      End
      Begin VB.ListBox lstReqItem 
         Height          =   1860
         ItemData        =   "frmEditor_Quest.frx":08ED
         Left            =   2760
         List            =   "frmEditor_Quest.frx":08EF
         TabIndex        =   79
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdReqItemRemove 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   4560
         TabIndex        =   77
         Top             =   3000
         Width           =   1215
      End
      Begin VB.HScrollBar scrlReqLevel 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   21
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar scrlReqQuest 
         Height          =   255
         Left            =   120
         Max             =   70
         TabIndex        =   20
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton cmdReqItem 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   2760
         TabIndex        =   78
         Top             =   3000
         Width           =   1815
      End
      Begin VB.CommandButton cmdReqClass 
         Caption         =   "Actualizar"
         Height          =   255
         Left            =   120
         TabIndex        =   85
         Top             =   3960
         Width           =   1215
      End
      Begin VB.Label lblReqSwitch 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Switch Req: Ninguno"
         Height          =   180
         Left            =   120
         TabIndex        =   89
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label lblReqClass 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Clase Req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   83
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label lblReqItem 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Objeto Req: 0 (1)"
         Height          =   180
         Left            =   2760
         TabIndex        =   82
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblReqLevel 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nivel Req: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblReqQuest 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mision Req: Ninguna"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////
'///////////////// QUEST SYSTEM - Developed by Alatar ////////////////
'/////////////////////////////////////////////////////////////////////

Option Explicit
Private TempTask As Long

Private Sub cmbSkill_Click()
    If cmbSkill.ListIndex < 0 Then Exit Sub
    If EditorIndex < 1 Then Exit Sub
    Quest(EditorIndex).Skill = cmbSkill.ListIndex + 1
End Sub

Private Sub cmdSSave_Click()
    If Options.Debug Then On Error GoTo ErrHandler
    
    If LenB(Trim$(txtName)) = 0 Then
        Call Msgbox("Nombre Requerido.")
    Else
        QuestEditorOk False
    End If
    
    Exit Sub
ErrHandler:
    HandleError "cmdSSave", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Exit Sub
End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L1")
Frame3.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L2")
Frame7.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L3")
optShowFrame(0).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L4")
optShowFrame(1).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L5")
optShowFrame(2).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L6")
optShowFrame(3).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L7")
fraTasks.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L8")
lblSelected.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L9")
lblSpeech.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L10")
lblLog.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L12")
lblItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L13")
lblMap.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L14")
lblResource.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L15")
lblEvent.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L16")
lblAmount.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L17")
chkEnd.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L18")
optTask(0).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L19")
optTask(1).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L20")
optTask(2).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L21")
optTask(3).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L22")
optTask(4).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L23")
optTask(5).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L24")
optTask(6).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L25")
optTask(7).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L26")
optTask(8).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L27")
optTask(9).Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L28")
fraRewards.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L29")
lblItemRew.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L30")
lblExp.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L31")
Label3.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L32")
lblSkillExp.Caption = trad

trad = GetVar(App.Path & Lang, "MissionSystem", "L33")
fraRequirements.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L34")
lblReqLevel.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L35")
lblReqItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L36")
lblReqQuest.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L37")
lblReqSwitch.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L38")
lblReqClass.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L39")
fraGeneral.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L40")
Label2.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L41")
chkRepeat.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L42")
lblQ1.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L43")
lblQ2.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L44")
lblQ3.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L45")
lblGiveItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "L46")
lblTakeItem.Caption = trad


trad = GetVar(App.Path & Lang, "MissionSystem", "BA")
cmdItemRew.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "BB")
cmdItemRewRemove.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B1")
cmdArray.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B2")
cmdSave.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B3")
cmdSSave.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B4")
cmdDelete.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B5")
cmdCancel.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B6")
cmdReqClass.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B7")
cmdReqClassRemove.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B8")
cmdReqItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B9")
cmdReqItemRemove.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B10")
cmdGiveItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B11")
cmdGiveItemRemove.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B12")
cmdTakeItem.Caption = trad
trad = GetVar(App.Path & Lang, "MissionSystem", "B13")
cmdTakeItemRemove.Caption = trad

'-------------------------------------------
    scrlTotalTasks.max = MAX_TASKS
    scrlNPC.max = MAX_NPCS
    scrlItem.max = MAX_ITEMS
    scrlMap.max = MAX_MAPS
    scrlResource.max = MAX_RESOURCES
    scrlAmount.max = MAX_INTEGER
    scrlReqLevel.max = MAX_LEVELS
    scrlReqQuest.max = MAX_QUESTS
    scrlReqItem.max = MAX_ITEMS
    scrlReqItemValue.max = MAX_INTEGER
    scrlGiveItem.max = MAX_ITEMS
    scrlGiveItemValue.max = MAX_INTEGER
    scrlTakeItem.max = MAX_ITEMS
    scrlTakeItemValue.max = MAX_INTEGER
    scrlExp.max = MAX_INTEGER 'Alatar v1.2
    scrlItemRew.max = MAX_ITEMS
    scrlItemRewValue.max = MAX_INTEGER
End Sub

Private Sub cmdSave_Click()
    If LenB(Trim$(txtName)) = 0 Then
        Call Msgbox("Nombre Requerido.")
    Else
        QuestEditorOk
    End If
End Sub

Private Sub cmdCancel_Click()
    QuestEditorCancel
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
End Sub

Private Sub scrlSkillExp_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L32")
    
    lblSkillExp.Caption = trad & " " & scrlSkillExp.Value
    Quest(EditorIndex).SkillExp = scrlSkillExp.Value
End Sub

Private Sub lstIndex_Click()
    QuestEditorInit
End Sub

Private Sub scrlEvent_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L15")
    
    If scrlEvent.Value > 0 Then
        lblEvent.Caption = trad & " " & scrlEvent.Value '& "-" & Map.Events(scrlEvent.Value).Name
    Else
        lblEvent.Caption = trad & " None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Event = scrlEvent.Value
End Sub

Private Sub scrlTotalTasks_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L8")
    
    Dim I As Long
    
    lblSelected = trad & " " & scrlTotalTasks.Value
    
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub optTask_Click(Index As Integer)
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Order = Index
    LoadTask EditorIndex, scrlTotalTasks.Value
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long
    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
End Sub

Private Sub txtQuestLog_Change()
    Quest(EditorIndex).QuestLog = Trim$(txtQuestLog.text)
End Sub

Private Sub txtTaskLog_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).TaskLog = Trim$(txtTaskLog.text)
End Sub

Private Sub chkRepeat_Click()
    If chkRepeat.Value = 1 Then
        Quest(EditorIndex).Repeat = 1
    Else
        Quest(EditorIndex).Repeat = 0
    End If
End Sub

Private Sub scrlReqLevel_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L34")
    
    lblReqLevel.Caption = trad & " " & scrlReqLevel.Value
    Quest(EditorIndex).RequiredLevel = scrlReqLevel.Value
End Sub

Private Sub scrlReqQuest_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L36")
    If Not scrlReqQuest.Value = 0 Then
        If Not Trim$(Quest(scrlReqQuest.Value).name) = "" Then
            lblReqQuest.Caption = trad & " " & Trim$(Quest(scrlReqQuest.Value).name)
        Else
            lblReqQuest.Caption = trad & " None"
        End If
    Else
        lblReqQuest.Caption = trad & " None"
    End If
    Quest(EditorIndex).RequiredQuest = scrlReqQuest.Value
End Sub

'Alatar v1.2

Private Sub scrlReqItem_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L35")
    
    If scrlReqItem.Value > 0 Then
        lblReqItem.Caption = trad & " " & Trim$(Item(scrlReqItem.Value).name) & " (" & scrlReqItemValue.Value & ")"
    Else
        scrlReqItemValue.Value = 1
        lblReqItem.Caption = trad & " None (" & scrlReqItemValue.Value & ")"
    End If
End Sub

Private Sub scrlReqItemValue_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L35")
    
    If scrlReqItem.Value > 0 Then lblReqItem.Caption = trad & " " & Trim$(Item(scrlReqItem.Value).name) & " (" & scrlReqItemValue.Value & ")"
End Sub

Private Sub cmdReqItem_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlReqItem.Value < 1 Or scrlReqItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlReqItem.Value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = scrlReqItem.Value
    Quest(EditorIndex).RequiredItem(Index).Value = scrlReqItemValue.Value
    UpdateQuestRequirementItems
End Sub

Private Sub cmdReqItemRemove_Click()
    Dim Index As Long
    
    Index = lstReqItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredItem(Index).Item = 0
    Quest(EditorIndex).RequiredItem(Index).Value = 1
    UpdateQuestRequirementItems
End Sub

Private Sub scrlReqClass_Change()
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then
        lblReqClass.Caption = "Clase: 0"
    Else
        lblReqClass.Caption = "Clase: " & scrlReqClass.Value & " (" & Trim$(Class(scrlReqClass.Value).name) & ")"
    End If
End Sub

Private Sub cmdReqClass_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1 'the selected class
    If Index = 0 Then Exit Sub
    If scrlReqClass.Value < 1 Or scrlReqClass.Value > Max_Classes Then Exit Sub
    If Trim$(Class(scrlReqClass.Value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = scrlReqClass.Value
    UpdateQuestClass
End Sub

Private Sub cmdReqClassRemove_Click()
    Dim Index As Long
    
    Index = lstReqClass.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RequiredClass(Index) = 0
    UpdateQuestClass
End Sub

'/Alatar v1.2

Private Sub scrlExp_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L30")
    
    lblExp = trad & " " & scrlExp.Value
    Quest(EditorIndex).RewardExp = scrlExp.Value
End Sub

Private Sub scrlItemRew_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L29")
    
    If scrlItemRew.Value > 0 Then
        lblItemRew.Caption = trad & " " & Trim$(Item(scrlItemRew.Value).name) & " (" & scrlItemRewValue.Value & ")"
    Else
        lblItemRew.Caption = trad & " None (" & scrlItemRewValue.Value & ")"
    End If
End Sub

Private Sub scrlItemRewValue_Change()
    If scrlItemRew.Value > 0 Then lblItemRew.Caption = "Objeto Recompensa: " & Trim$(Item(scrlItemRew.Value).name) & " (" & scrlItemRewValue.Value & ")"
End Sub

'Alatar v1.2
Private Sub cmdItemRew_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlItemRew.Value < 1 Or scrlItemRew.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlItemRew.Value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = scrlItemRew.Value
    Quest(EditorIndex).RewardItem(Index).Value = scrlItemRewValue.Value
    UpdateQuestRewardItems
End Sub

Private Sub cmdItemRewRemove_Click()
    Dim Index As Long
    
    Index = lstItemRew.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).RewardItem(Index).Item = 0
    Quest(EditorIndex).RewardItem(Index).Value = 1
    UpdateQuestRewardItems
End Sub
'/Alatar v1.2

Private Sub txtSpeech_Change(Index As Integer)
    Quest(EditorIndex).Speech(Index) = Trim$(txtSpeech(Index).text)
End Sub

Private Sub txtTaskSpeech_Change()
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Speech = Trim$(txtTaskSpeech.text)
End Sub

'Alatar v1.2
Private Sub scrlGiveItem_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L45")
    
    If scrlGiveItem.Value > 0 Then
        lblGiveItem = trad & " " & Trim$(Item(scrlGiveItem.Value).name) & " (" & scrlGiveItemValue.Value & ")"
    Else
        scrlGiveItemValue.Value = 1
        lblGiveItem = trad & " None (" & scrlGiveItemValue.Value & ")"
    End If
End Sub

Private Sub scrlGiveItemValue_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L45")
    
    If scrlGiveItem.Value > 0 Then lblGiveItem = trad & " " & Trim$(Item(scrlGiveItem.Value).name) & " (" & scrlGiveItemValue.Value & ")"
End Sub

Private Sub cmdGiveItem_Click()
    Dim Index As Long
    
    Index = lstGiveItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlGiveItem.Value < 1 Or scrlGiveItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlGiveItem.Value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).GiveItem(Index).Item = scrlGiveItem.Value
    Quest(EditorIndex).GiveItem(Index).Value = scrlGiveItemValue.Value
    UpdateQuestGiveItems
End Sub

Private Sub cmdGiveItemRemove_Click()
    Dim Index As Long
    
    Index = lstGiveItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).GiveItem(Index).Item = 0
    Quest(EditorIndex).GiveItem(Index).Value = 1
    UpdateQuestGiveItems
End Sub

Private Sub scrlTakeItem_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L46")
    
    If scrlTakeItem.Value > 0 Then
        lblTakeItem = trad & " " & Trim$(Item(scrlTakeItem.Value).name) & " (" & scrlTakeItemValue.Value & ")"
    Else
        scrlTakeItemValue.Value = 1
        lblTakeItem = trad & " None (" & scrlTakeItemValue.Value & ")"
    End If
End Sub

Private Sub scrlTakeItemValue_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L46")
    
    If scrlTakeItem.Value > 0 Then lblTakeItem = trad & " " & Trim$(Item(scrlTakeItem.Value).name) & " (" & scrlTakeItemValue.Value & ")"
End Sub

Private Sub cmdTakeItem_Click()
    Dim Index As Long
    
    Index = lstTakeItem.ListIndex + 1 'the selected item
    If Index = 0 Then Exit Sub
    If scrlTakeItem.Value < 1 Or scrlTakeItem.Value > MAX_ITEMS Then Exit Sub
    If Trim$(Item(scrlTakeItem.Value).name) = "" Then Exit Sub
    
    Quest(EditorIndex).TakeItem(Index).Item = scrlTakeItem.Value
    Quest(EditorIndex).TakeItem(Index).Value = scrlTakeItemValue.Value
    UpdateQuestTakeItems
End Sub

Private Sub cmdTakeItemRemove_Click()
    Dim Index As Long
    
    Index = lstTakeItem.ListIndex + 1
    If Index = 0 Then Exit Sub
    
    Quest(EditorIndex).TakeItem(Index).Item = 0
    Quest(EditorIndex).TakeItem(Index).Value = 1
    UpdateQuestTakeItems
End Sub
'/Alatar v1.2

Private Sub scrlAmount_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L16")
    lblAmount.Caption = trad & " " & scrlAmount.Value
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Amount = scrlAmount.Value
End Sub

Private Sub scrlNPC_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L11")
    
    If scrlNPC.Value > 0 Then
        lblNPC.Caption = trad & " " & scrlNPC.Value & "-" & Trim$(NPC(scrlNPC.Value).name)
    Else
        lblNPC.Caption = trad & " None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).NPC = scrlNPC.Value
End Sub

Private Sub scrlItem_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L12")
    
    If scrlItem.Value > 0 Then
        lblItem.Caption = trad & " " & Trim$(Item(scrlItem.Value).name)
    Else
        lblItem.Caption = trad & " None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Item = scrlItem.Value
End Sub

Private Sub scrlMap_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L13")
    
    If scrlMap.Value > 0 Then
        lblMap.Caption = trad & " " & scrlMap.Value
    Else
        lblMap.Caption = trad & " None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Map = scrlMap.Value
End Sub

Private Sub scrlResource_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "MissionSystem", "L14")
    
    If scrlResource.Value > 0 Then
        lblResource.Caption = trad & " " & scrlResource.Value & "-" & Trim$(Resource(scrlResource.Value).name)
    Else
        lblResource.Caption = trad & " None"
    End If
    Quest(EditorIndex).Task(scrlTotalTasks.Value).Resource = scrlResource.Value
End Sub

Private Sub chkEnd_Click()
    If chkEnd.Value = 1 Then
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = True
    Else
        Quest(EditorIndex).Task(scrlTotalTasks.Value).QuestEnd = False
    End If
End Sub

Private Sub optShowFrame_Click(Index As Integer)
    fraGeneral.Visible = False
    fraRequirements.Visible = False
    fraRewards.Visible = False
    fraTasks.Visible = False
    
    If optShowFrame(Index).Value = True Then
        Select Case Index
            Case 0
                fraGeneral.Visible = True
            Case 1
                fraRequirements.Visible = True
            Case 2
                fraRewards.Visible = True
            Case 3
                fraTasks.Visible = True
        End Select
    End If
End Sub
