VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCN.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form frmServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargando..."
   ClientHeight    =   6795
   ClientLeft      =   1440
   ClientTop       =   2535
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmServer.frx":4FEA
   ScaleHeight     =   6795
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "MODO BASICO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   225
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H00808000&
      Height          =   3615
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   224
      Top             =   1560
      Width           =   7155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   0
   End
   Begin VB.PictureBox updatebtn 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   10720
      Picture         =   "frmServer.frx":1E614
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   190
      Top             =   5760
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Timer tmrGetTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Cmdboton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1050
      Index           =   3
      Left            =   1200
      Picture         =   "frmServer.frx":2043E
      ScaleHeight     =   1050
      ScaleWidth      =   1200
      TabIndex        =   3
      ToolTipText     =   "Editors Machine. Paneles de Edicion de EaSee para editar clases, clanes, maximos, muerte, clima, etc, afectando la jugabilidad."
      Top             =   4680
      Width           =   1200
   End
   Begin VB.PictureBox Cmdboton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1050
      Index           =   2
      Left            =   1320
      Picture         =   "frmServer.frx":23019
      ScaleHeight     =   1050
      ScaleWidth      =   1200
      TabIndex        =   2
      ToolTipText     =   "Configuracion Principal. Recarga mapas, objetos, etc. Configuraciones extra del servidor."
      Top             =   3600
      Width           =   1200
   End
   Begin VB.PictureBox Cmdboton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   885
      Index           =   1
      Left            =   1440
      Picture         =   "frmServer.frx":25D57
      ScaleHeight     =   885
      ScaleWidth      =   1200
      TabIndex        =   1
      ToolTipText     =   $"frmServer.frx":28CE3
      Top             =   2640
      Width           =   1200
   End
   Begin VB.PictureBox Cmdboton 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   1050
      Index           =   0
      Left            =   1680
      Picture         =   "frmServer.frx":28D83
      ScaleHeight     =   1050
      ScaleWidth      =   1200
      TabIndex        =   0
      ToolTipText     =   "Consola General. Estado, Logs del Server y Chat Administrativo."
      Top             =   1440
      Width           =   1200
   End
   Begin VB.PictureBox cmdShutDown 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2880
      Picture         =   "frmServer.frx":2BAF9
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   13
      ToolTipText     =   "Apagar Servidor"
      Top             =   5835
      Width           =   600
   End
   Begin VB.PictureBox Picconsola 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3600
      Picture         =   "frmServer.frx":2DF2A
      ScaleHeight     =   4335
      ScaleWidth      =   7575
      TabIndex        =   4
      Top             =   1320
      Width           =   7575
      Begin VB.TextBox txtChat 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   4040
         Width           =   7260
      End
   End
   Begin VB.PictureBox PicEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3600
      Picture         =   "frmServer.frx":32469
      ScaleHeight     =   4335
      ScaleWidth      =   7575
      TabIndex        =   45
      Top             =   1320
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox CmdEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   4
         Left            =   5400
         Picture         =   "frmServer.frx":36392
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   189
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox CmdEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   3
         Left            =   4080
         Picture         =   "frmServer.frx":39A6E
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   49
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox CmdEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   1
         Left            =   1440
         Picture         =   "frmServer.frx":3D21C
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   48
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox CmdEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   0
         Left            =   120
         Picture         =   "frmServer.frx":409DE
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   47
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox CmdEdit 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Index           =   2
         Left            =   2760
         Picture         =   "frmServer.frx":442B3
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   46
         Top             =   120
         Width           =   1200
      End
      Begin VB.PictureBox PicMaximos 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   4215
         TabIndex        =   122
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
         Begin VB.PictureBox txtguardarmaxmin 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   1560
            Picture         =   "frmServer.frx":47731
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   125
            Top             =   2160
            Width           =   1050
         End
         Begin VB.TextBox txtmaxclases 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            MaxLength       =   2
            TabIndex        =   124
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Clases:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.PictureBox PicClima 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   4215
         TabIndex        =   107
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
         Begin VB.PictureBox Pricmaquina 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   2880
            Picture         =   "frmServer.frx":49B69
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   121
            Top             =   1680
            Width           =   1050
         End
         Begin VB.PictureBox cmdclima 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   2880
            Picture         =   "frmServer.frx":4C376
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   120
            Top             =   2280
            Width           =   1050
         End
         Begin VB.TextBox txtopacidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   114
            Text            =   "255"
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtvelocidad 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   113
            Text            =   "255"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtniebla 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   112
            Text            =   "255"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtintensidadclima 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   111
            Text            =   "100"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cmbclima 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":4EAFF
            Left            =   120
            List            =   "frmServer.frx":4EB15
            TabIndex        =   110
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtmapaclima 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   108
            Text            =   "0"
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lbltipoclima 
            BackStyle       =   0  'Transparent
            Caption         =   "Clima"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   119
            Top             =   640
            Width           =   495
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Intensidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   118
            Top             =   1120
            Width           =   975
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Niebla"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   117
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Velocidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   116
            Top             =   2080
            Width           =   855
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Opacidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1440
            TabIndex        =   115
            Top             =   2560
            Width           =   855
         End
         Begin VB.Label lblmapa 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero Mapa (0 para Todos)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   1440
            TabIndex        =   109
            Top             =   140
            Width           =   2520
         End
      End
      Begin VB.PictureBox PicMaquinaClima 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   7335
         TabIndex        =   50
         Top             =   1200
         Visible         =   0   'False
         Width           =   7335
         Begin VB.PictureBox CmdAtrasClima 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   6000
            Picture         =   "frmServer.frx":4EB52
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   106
            Top             =   2520
            Width           =   1050
         End
         Begin VB.PictureBox checkclima 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   5760
            Picture         =   "frmServer.frx":51257
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   103
            Top             =   2040
            Width           =   255
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   6480
            TabIndex        =   51
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
            Begin VB.CheckBox chkclimas 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Desact Maquina"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   53
               Top             =   720
               Width           =   1935
            End
            Begin VB.CheckBox chkrandclima 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Clima Aleatorio"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Value           =   1  'Checked
               Width           =   1695
            End
         End
         Begin VB.PictureBox checkclima 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   5760
            Picture         =   "frmServer.frx":53119
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   102
            Top             =   1680
            Width           =   255
         End
         Begin VB.ComboBox cmbhora4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":54FDB
            Left            =   720
            List            =   "frmServer.frx":54FF1
            TabIndex        =   77
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":5502E
            Left            =   720
            List            =   "frmServer.frx":55044
            TabIndex        =   76
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55081
            Left            =   720
            List            =   "frmServer.frx":55097
            TabIndex        =   75
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":550D4
            Left            =   720
            List            =   "frmServer.frx":550EA
            TabIndex        =   74
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora0 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55127
            Left            =   720
            List            =   "frmServer.frx":5513D
            TabIndex        =   73
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora6 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":5517A
            Left            =   720
            List            =   "frmServer.frx":55190
            TabIndex        =   72
            Top             =   2520
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":551CD
            Left            =   720
            List            =   "frmServer.frx":551E3
            TabIndex        =   71
            Top             =   2160
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora9 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55220
            Left            =   2640
            List            =   "frmServer.frx":55236
            TabIndex        =   70
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora7 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55273
            Left            =   2640
            List            =   "frmServer.frx":55289
            TabIndex        =   69
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora8 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":552C6
            Left            =   2640
            List            =   "frmServer.frx":552DC
            TabIndex        =   68
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora10 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55319
            Left            =   2640
            List            =   "frmServer.frx":5532F
            TabIndex        =   67
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora12 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":5536C
            Left            =   2640
            List            =   "frmServer.frx":55382
            TabIndex        =   66
            Top             =   2160
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora13 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":553BF
            Left            =   2640
            List            =   "frmServer.frx":553D5
            TabIndex        =   65
            Top             =   2520
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora11 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55412
            Left            =   2640
            List            =   "frmServer.frx":55428
            TabIndex        =   64
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora15 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55465
            Left            =   4440
            List            =   "frmServer.frx":5547B
            TabIndex        =   63
            Top             =   720
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora14 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":554B8
            Left            =   4440
            List            =   "frmServer.frx":554CE
            TabIndex        =   62
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora20 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":5550B
            Left            =   4440
            List            =   "frmServer.frx":55521
            TabIndex        =   61
            Top             =   2520
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora17 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":5555E
            Left            =   4440
            List            =   "frmServer.frx":55574
            TabIndex        =   60
            Top             =   1440
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora18 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":555B1
            Left            =   4440
            List            =   "frmServer.frx":555C7
            TabIndex        =   59
            Top             =   1800
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora19 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55604
            Left            =   4440
            List            =   "frmServer.frx":5561A
            TabIndex        =   58
            Top             =   2160
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora16 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55657
            Left            =   4440
            List            =   "frmServer.frx":5566D
            TabIndex        =   57
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora21 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":556AA
            Left            =   6240
            List            =   "frmServer.frx":556C0
            TabIndex        =   56
            Top             =   360
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora23 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":556FD
            Left            =   6240
            List            =   "frmServer.frx":55713
            TabIndex        =   55
            Top             =   1080
            Width           =   1095
         End
         Begin VB.ComboBox cmbhora22 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmServer.frx":55750
            Left            =   6240
            List            =   "frmServer.frx":55766
            TabIndex        =   54
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Lblcheckclima 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aleatorio"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   6120
            TabIndex        =   105
            Top             =   1720
            Width           =   765
         End
         Begin VB.Label Lblcheckclima 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Maquina"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   0
            Left            =   6120
            TabIndex        =   104
            Top             =   2080
            Width           =   990
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "4 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Top             =   1830
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "3 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   100
            Top             =   1455
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "2 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   99
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "1 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   98
            Top             =   735
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "0 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "6 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   96
            Top             =   2550
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "5 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   95
            Top             =   2200
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "7 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   94
            Top             =   420
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "8 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   8
            Left            =   2040
            TabIndex        =   93
            Top             =   765
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "10 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   10
            Left            =   1920
            TabIndex        =   92
            Top             =   1515
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "9 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   11
            Left            =   2040
            TabIndex        =   91
            Top             =   1125
            Width           =   615
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "1 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   13
            Left            =   2040
            TabIndex        =   90
            Top             =   2550
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "12 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   14
            Left            =   1920
            TabIndex        =   89
            Top             =   2205
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "11 Am"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   15
            Left            =   1920
            TabIndex        =   88
            Top             =   1830
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "3 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   9
            Left            =   3840
            TabIndex        =   87
            Top             =   780
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "2 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   12
            Left            =   3840
            TabIndex        =   86
            Top             =   405
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "6 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   16
            Left            =   3840
            TabIndex        =   85
            Top             =   1860
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "5 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   17
            Left            =   3840
            TabIndex        =   84
            Top             =   1485
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "4 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   18
            Left            =   3840
            TabIndex        =   83
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "7 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   19
            Left            =   3840
            TabIndex        =   82
            Top             =   2205
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "8 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   21
            Left            =   3840
            TabIndex        =   81
            Top             =   2565
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "9 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   20
            Left            =   5760
            TabIndex        =   80
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "11 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   22
            Left            =   5640
            TabIndex        =   79
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label lblhora 
            BackStyle       =   0  'Transparent
            Caption         =   "10 Pm"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   23
            Left            =   5640
            TabIndex        =   78
            Top             =   770
            Width           =   735
         End
      End
      Begin VB.PictureBox PicEditClases 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         ScaleHeight     =   3015
         ScaleWidth      =   7335
         TabIndex        =   154
         Top             =   1200
         Visible         =   0   'False
         Width           =   7335
         Begin VB.TextBox txtevolclase 
            Height          =   315
            Left            =   2760
            TabIndex        =   7
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtevolucionniv 
            Height          =   315
            Left            =   2760
            TabIndex        =   212
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtitemfaccion 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            MaxLength       =   3
            TabIndex        =   209
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox txtItemC1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   206
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtItemC2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            MaxLength       =   2
            TabIndex        =   205
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtItemC3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   204
            Text            =   "0"
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtItemN1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   203
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtItemN2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            MaxLength       =   2
            TabIndex        =   202
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtItemN3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   201
            Text            =   "0"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox txtSpellN1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   199
            Text            =   "0"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtSpellN2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            MaxLength       =   2
            TabIndex        =   198
            Text            =   "0"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtSpellN3 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6600
            MaxLength       =   2
            TabIndex        =   197
            Text            =   "0"
            Top             =   480
            Width           =   495
         End
         Begin VB.HScrollBar hscrollfaccion 
            Height          =   255
            Left            =   1320
            Max             =   4
            Min             =   1
            TabIndex        =   192
            Top             =   450
            Value           =   1
            Width           =   2175
         End
         Begin VB.PictureBox CheckVisibleClase 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   3600
            Picture         =   "frmServer.frx":557A3
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   188
            Top             =   120
            Width           =   255
         End
         Begin VB.PictureBox cmdguardarclases 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   6120
            Picture         =   "frmServer.frx":57665
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   186
            Top             =   2520
            Width           =   1050
         End
         Begin VB.TextBox txtcorrervelocidad 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   185
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtcaminarvelocidad 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   184
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtyspawn 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5640
            TabIndex        =   180
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtxspawn 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            TabIndex        =   178
            Top             =   2640
            Width           =   375
         End
         Begin VB.TextBox txtmapaspawn 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   176
            Top             =   2640
            Width           =   495
         End
         Begin VB.CheckBox chkvisible 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oculta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   7200
            TabIndex        =   175
            Top             =   2400
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtvoluntadclase 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   172
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtagilidadclase 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   171
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtresistenciaclase 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   170
            Top             =   2640
            Width           =   495
         End
         Begin VB.TextBox txtinteligenciaclase 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   169
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox txtfuerzaclase 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   168
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtspritefemclase 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6120
            TabIndex        =   161
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox txtspritemascclase 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6120
            TabIndex        =   160
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtnombreclase 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   158
            Top             =   800
            Width           =   2535
         End
         Begin VB.HScrollBar sldclasenum 
            Height          =   255
            Left            =   1320
            Min             =   1
            TabIndex        =   156
            Top             =   90
            Value           =   1
            Width           =   2175
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   0
            TabIndex        =   223
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "Facción:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Left            =   0
            TabIndex        =   222
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Clase:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   2040
            TabIndex        =   214
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   2040
            TabIndex        =   213
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblsprite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprites Femenino:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   1
            Left            =   4440
            TabIndex        =   162
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dropear Item:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   5280
            TabIndex        =   210
            Top             =   150
            Width           =   1245
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1:           2:           3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   2
            Left            =   4680
            TabIndex        =   208
            Top             =   1240
            Width           =   1860
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1:           2:           3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   1
            Left            =   4680
            TabIndex        =   207
            Top             =   870
            Width           =   1860
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1:           2:           3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   4680
            TabIndex        =   200
            Top             =   525
            Width           =   1860
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   3720
            TabIndex        =   196
            Top             =   1240
            Width           =   840
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   3960
            TabIndex        =   195
            Top             =   870
            Width           =   570
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hechizos:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   3720
            TabIndex        =   194
            Top             =   510
            Width           =   825
         End
         Begin VB.Label LblFacionNumero 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   960
            TabIndex        =   193
            Top             =   450
            Width           =   105
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clase Oculta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   3900
            TabIndex        =   187
            Top             =   150
            Width           =   1080
         End
         Begin VB.Label lblcorrer 
            BackStyle       =   0  'Transparent
            Caption         =   "Correr:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   2160
            TabIndex        =   183
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label lblcaminar 
            BackStyle       =   0  'Transparent
            Caption         =   "Caminar:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   1920
            TabIndex        =   182
            Top             =   1560
            Width           =   855
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
            ForeColor       =   &H00C0C000&
            Height          =   375
            Left            =   2160
            TabIndex        =   181
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblyspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Y:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   5400
            TabIndex        =   179
            Top             =   2670
            Width           =   255
         End
         Begin VB.Label lblxspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "X:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   4680
            TabIndex        =   177
            Top             =   2670
            Width           =   255
         End
         Begin VB.Label lblmapaspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Mapa:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   3480
            TabIndex        =   174
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblspawn 
            BackStyle       =   0  'Transparent
            Caption         =   "Spawn:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   375
            Left            =   3480
            TabIndex        =   173
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblStats 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Voluntad:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   167
            Top             =   1950
            Width           =   1215
         End
         Begin VB.Label lblStats 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agilidad:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   3
            Left            =   0
            TabIndex        =   166
            Top             =   1590
            Width           =   1230
         End
         Begin VB.Label lblStats 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resistencia:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   2
            Left            =   0
            TabIndex        =   165
            Top             =   2670
            Width           =   1155
         End
         Begin VB.Label lblStats 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inteligencia:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   164
            Top             =   2310
            Width           =   1185
         End
         Begin VB.Label lblStats 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuerza:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   163
            Top             =   1230
            Width           =   1245
         End
         Begin VB.Label lblsprite 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sprites Masculino:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   4440
            TabIndex        =   159
            Top             =   1710
            Width           =   1560
         End
         Begin VB.Label lblnombre 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   157
            Top             =   810
            Width           =   735
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   960
            TabIndex        =   155
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.PictureBox PicGremio 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   5415
         TabIndex        =   138
         Top             =   1320
         Visible         =   0   'False
         Width           =   5415
         Begin VB.TextBox txtGJoinLvl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   152
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtGJoinItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   149
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtGJoinCost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4080
            TabIndex        =   148
            Top             =   1080
            Width           =   975
         End
         Begin VB.PictureBox cmdGSave 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   450
            Left            =   2160
            Picture         =   "frmServer.frx":59A9D
            ScaleHeight     =   450
            ScaleWidth      =   1050
            TabIndex        =   147
            Top             =   2280
            Width           =   1050
         End
         Begin VB.TextBox txtGBuyLvl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   143
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtGBuyCost 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   142
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtGBuyItem 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   141
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Req:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   3120
            TabIndex        =   153
            Top             =   1590
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   3120
            TabIndex        =   151
            Top             =   1110
            Width           =   840
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   3240
            TabIndex        =   150
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Objeto:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   480
            TabIndex        =   146
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nivel Req:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   240
            TabIndex        =   145
            Top             =   1590
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Left            =   360
            TabIndex        =   144
            Top             =   1110
            Width           =   840
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Coste de Nuevo miembro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   2760
            TabIndex        =   140
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Coste de creacion."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   120
            Width           =   1815
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00808000&
            Height          =   1695
            Left            =   2760
            Top             =   360
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00808000&
            Height          =   1695
            Left            =   120
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.PictureBox PicMenumuerte 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   2895
         Left            =   120
         ScaleHeight     =   2895
         ScaleWidth      =   3855
         TabIndex        =   126
         Top             =   1320
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Muerte"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   960
            TabIndex        =   133
            Top             =   600
            Visible         =   0   'False
            Width           =   2535
            Begin VB.CheckBox chkdropmuerte 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2280
               TabIndex        =   135
               Top             =   720
               Width           =   255
            End
            Begin VB.CheckBox chkbloqpj 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2280
               TabIndex        =   134
               Top             =   360
               Width           =   255
            End
            Begin VB.Label lblmuerte1 
               BackStyle       =   0  'Transparent
               Caption         =   "Perder Items al Morir"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   720
               Width           =   2055
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Borrar Cuenta al Morir"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   136
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.TextBox txtprobabilidaddrop 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1680
            MaxLength       =   3
            TabIndex        =   132
            Text            =   "100"
            Top             =   1440
            Width           =   495
         End
         Begin VB.PictureBox ChckMenuMuerte 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   360
            Picture         =   "frmServer.frx":5BED5
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   128
            Top             =   840
            Width           =   255
         End
         Begin VB.PictureBox ChckMenuMuerte 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   360
            Picture         =   "frmServer.frx":5DD97
            ScaleHeight     =   270
            ScaleWidth      =   255
            TabIndex        =   127
            Top             =   480
            Width           =   255
         End
         Begin VB.Label LblMenumorir 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Probabilidades"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   131
            Top             =   1440
            Width           =   1245
         End
         Begin VB.Label LblMenumorir 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Perder objetos al morir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   1
            Left            =   720
            TabIndex        =   130
            Top             =   870
            Width           =   1995
         End
         Begin VB.Label LblMenumorir 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Borrar cuenta al morir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C000&
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   129
            Top             =   525
            Width           =   1920
         End
      End
   End
   Begin VB.PictureBox PicControl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
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
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3600
      Picture         =   "frmServer.frx":5FC59
      ScaleHeight     =   4335
      ScaleWidth      =   7575
      TabIndex        =   9
      Top             =   1320
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   4320
         Picture         =   "frmServer.frx":65260
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   219
         Top             =   2280
         Width           =   255
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   4320
         Picture         =   "frmServer.frx":67122
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   217
         Top             =   1920
         Width           =   255
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   4320
         Picture         =   "frmServer.frx":68FE4
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   42
         Top             =   2640
         Width           =   255
      End
      Begin VB.PictureBox btnDubExp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   5455
         Picture         =   "frmServer.frx":6AEA6
         ScaleHeight     =   1050
         ScaleWidth      =   1875
         TabIndex        =   41
         Top             =   3100
         Width           =   1875
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   4320
         Picture         =   "frmServer.frx":6E758
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   36
         Top             =   1560
         Width           =   255
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   4320
         Picture         =   "frmServer.frx":7061A
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   35
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   4320
         Picture         =   "frmServer.frx":724DC
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   33
         Top             =   840
         Width           =   255
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4320
         Picture         =   "frmServer.frx":7439E
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   32
         Top             =   480
         Width           =   255
      End
      Begin VB.PictureBox cmdReloadAnimations 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   3940
         Picture         =   "frmServer.frx":76260
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   24
         Top             =   3120
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadCombos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   2700
         Picture         =   "frmServer.frx":798C1
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   23
         Top             =   3120
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadShops 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   2700
         Picture         =   "frmServer.frx":7D2E3
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   22
         Top             =   2040
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadMaps 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   2700
         Picture         =   "frmServer.frx":80809
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   21
         Top             =   960
         Width           =   1200
      End
      Begin VB.PictureBox CmdReloadSpells 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   1470
         Picture         =   "frmServer.frx":8420C
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   20
         Top             =   3120
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadQuests 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   240
         Picture         =   "frmServer.frx":87916
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   19
         Top             =   3120
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadItems 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   240
         Picture         =   "frmServer.frx":8AE62
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   18
         Top             =   2040
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadClasses 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   240
         Picture         =   "frmServer.frx":8E259
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   17
         Top             =   960
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadResources 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   1470
         Picture         =   "frmServer.frx":91935
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   16
         Top             =   960
         Width           =   1200
      End
      Begin VB.PictureBox cmdReloadNPCs 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   1050
         Left            =   1470
         Picture         =   "frmServer.frx":94BA3
         ScaleHeight     =   1050
         ScaleWidth      =   1200
         TabIndex        =   15
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   6600
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
         Begin VB.CheckBox chkataliado 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ataque Aliado"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   221
            Top             =   2760
            Width           =   1455
         End
         Begin VB.CheckBox chk5frames 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Anim sprite"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   216
            Top             =   2520
            Width           =   1455
         End
         Begin VB.CheckBox chkServerLog 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Logs"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            MaskColor       =   &H00C0C0FF&
            TabIndex        =   44
            Top             =   2160
            Width           =   1575
         End
         Begin VB.CheckBox chkGUIBars 
            BackColor       =   &H00E0E0E0&
            Caption         =   "GUI Original(NO)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   1800
            Width           =   1815
         End
         Begin VB.CheckBox chkProj 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proyectiles (NO)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1440
            Width           =   4455
         End
         Begin VB.CheckBox chkFS 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Pantalla Completa (NO)"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1080
            Width           =   4455
         End
         Begin VB.CheckBox chkDropInvItems 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vaciar Inventario al Morir (Desact)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   4455
         End
         Begin VB.CheckBox chkFriendSystem 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sistema de Amistad (Desact)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.PictureBox Pickcheck 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3480
         Picture         =   "frmServer.frx":98106
         ScaleHeight     =   270
         ScaleWidth      =   255
         TabIndex        =   34
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Atacar Aliado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   220
         Top             =   2295
         Width           =   1140
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5 Anim sprite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   218
         Top             =   1950
         Width           =   1140
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOG (Registros)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   43
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de amistad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   4700
         TabIndex        =   37
         Top             =   510
         Width           =   1695
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GUI Original"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   40
         Top             =   1590
         Width           =   1050
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vaciar Inventario al morir"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   39
         Top             =   1230
         Width           =   2220
      End
      Begin VB.Label LblCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proyectiles"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   38
         Top             =   870
         Width           =   930
      End
   End
   Begin VB.PictureBox PicJugadores 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
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
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3600
      Picture         =   "frmServer.frx":99F66
      ScaleHeight     =   4335
      ScaleWidth      =   7575
      TabIndex        =   6
      Top             =   1320
      Visible         =   0   'False
      Width           =   7575
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   4215
         Left            =   50
         TabIndex        =   215
         Top             =   325
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   7435
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   12632064
         BackColor       =   0
         Appearance      =   0
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
            Object.Width           =   3821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cuenta"
            Object.Width           =   3821
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Personaje"
            Object.Width           =   3821
         EndProperty
         Picture         =   "frmServer.frx":A66B1
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº             Direccion de IP                     Cuenta                        Personaje "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   90
         Width           =   6525
      End
   End
   Begin VB.Label updatelbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.2"
      ForeColor       =   &H00FF00FF&
      Height          =   210
      Left            =   315
      TabIndex        =   211
      Top             =   6435
      Width           =   450
   End
   Begin VB.Label updatestate 
      BackStyle       =   0  'Transparent
      Caption         =   "EaSee Engine"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   191
      Top             =   5700
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Image ImgLogo 
      Height          =   495
      Left            =   840
      Top             =   550
      Width           =   1095
   End
   Begin VB.Label LblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BIENVENIDO, CLICK EN EL OJO PARA VISITARNOS EN NUESTRO SITIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2760
      TabIndex        =   25
      Top             =   400
      Width           =   6855
   End
   Begin VB.Label LblApagagando 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   3960
      TabIndex        =   14
      Top             =   5920
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblCpsLock 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[Desbloq]"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3720
      TabIndex        =   12
      Top             =   960
      Width           =   840
   End
   Begin VB.Label lblCPS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPS: 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4680
      TabIndex        =   11
      Top             =   960
      Width           =   600
   End
   Begin VB.Label lbltiempo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 00:00:00"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   480
      Left            =   6120
      TabIndex        =   10
      Top             =   6000
      Width           =   2025
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

Dim Contador As String

Option Explicit

Private Sub Command1_Click()
    'Player(1).Switches(1) = 1
    If Command1.Caption = "MODO BASICO" Then
    txtText.Top = 1560
    txtText.Left = 3840
    txtText.Height = 3615
    txtText.Width = 7155
    Command1.Caption = "MODO CONSOLA"
    Else
    txtText.Height = Me.Height - 400
    txtText.Width = Me.Width - 100
    txtText.Left = 0
    txtText.Top = 0
    Command1.Caption = "MODO BASICO "
    End If
End Sub

Private Sub Command2_Click()
    Player(1).Switches(1) = 0
End Sub
'spriteHOM = GetVar(App.Path & "\data\classes.ini", "CLASS" & GetPlayerClass(index), "MaleSprite")
Private Sub btnDubExp_Click()                                           '[EXP]
    Dim var1 As String

    DoubleExp = Not DoubleExp
    If DoubleExp Then
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "G1")
        Call GlobalMsg(var1, Green)          '[EXP]G1
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "T1")
        Call TextAdd(var1)                             '[EXP]T1
        btnDubExp.Picture = LoadPicture(App.Path & "\data\GUI\15_S.Jpg")
    Else
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "G2")
        Call GlobalMsg(var1, Green)       '[EXP]G2
        var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "T2")
        Call TextAdd(var1)                          '[EXP]T2
        btnDubExp.Picture = LoadPicture(App.Path & "\data\GUI\15.Jpg")
    End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub btnDubExp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

If DoubleExp Then                                                  '[EXP]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "C1")
        LblEstado.Caption = var1              '[EXP]C1
    Else
    var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EXP", "C2")
        LblEstado.Caption = var1                 '[EXP]C2
    End If
End Sub

Private Sub ChckMenuMuerte_Click(index As Integer)
Dim Path As String
    

Select Case index

         Case 0
         Path = App.Path & "\data\easee.ini"
              If chkbloqpj.Value = 1 Then
              chkbloqpj.Value = 0
              ChckMenuMuerte(0).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
              Else
              chkbloqpj.Value = 1
              ChckMenuMuerte(0).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
              End If
              Call PutVar(Path, "Jugador", "Muertebloq", CStr(chkbloqpj.Value))
         Case 1
          Path = App.Path & "\data\options.ini"
              If chkdropmuerte.Value = 1 Then
              chkdropmuerte.Value = 0
              ChckMenuMuerte(1).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
              Else
              chkdropmuerte.Value = 1
              ChckMenuMuerte(1).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
              End If
         Call PutVar(Path, "OPTIONS", "DropOnDeath", CStr(chkdropmuerte.Value))
End Select
         
         
         
End Sub

Private Sub ChckMenuMuerte_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

Select Case index

         Case 0                                                                 '[MUERTE]
              If chkbloqpj.Value Then
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "MUERTE", "C1")
              LblEstado.Caption = var1    '[MUERTE]C1
              Else
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "MUERTE", "C2")
              LblEstado.Caption = var1       '[MUERTE]C2
              End If
         Case 1
              If chkdropmuerte.Value Then
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "MUERTE", "C3")
              LblEstado.Caption = var1    '[MUERTE]C3
              Else
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "MUERTE", "C4")
              LblEstado.Caption = var1       '[MUERTE]C4
              End If
         
End Select
End Sub

Private Sub checkclima_Click(index As Integer)
Select Case index
          Case 0
               If chkrandclima.Value = 1 Then
               chkrandclima.Value = 0
               checkclima(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
               Else
               chkrandclima.Value = 1
               checkclima(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
               End If
          Case 1
               If chkclimas.Value = 1 Then
               chkclimas.Value = 0
               checkclima(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
               Else
               chkclimas.Value = 1
               checkclima(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
               End If
          
End Select
End Sub

Private Sub checkclima_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

Select Case index
          Case 0
               If chkrandclima.Value Then                           '[CLIMA]
               
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C1")
               LblEstado.Caption = var1        '[CLIMA]C1
               Else
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C2")
               LblEstado.Caption = var1    '[CLIMA]C2
               End If
          Case 1
               If chkclimas.Value Then
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C3")
               LblEstado.Caption = var1                '[CLIMA]C3
               Else
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C4")
               LblEstado.Caption = var1             '[CLIMA]C4
               End If
          
End Select
End Sub

Private Sub CheckVisibleClase_Click()
If chkvisible.Value = 0 Then
         chkvisible.Value = 1
         CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         Else
         chkvisible.Value = 0
         CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
End If
End Sub

Private Sub CheckVisibleClase_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

                                                '[CLASEVISIBLE]
                                                
            var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASEVISIBLE", "C1")
LblEstado.Caption = var1                  '[CLASEVISIBLE]C1
If chkvisible.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASEVISIBLE", "C2")
         LblEstado.Caption = var1    '[CLASEVISIBLE]C2
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASEVISIBLE", "C3")
         LblEstado.Caption = var1    '[CLASEVISIBLE]C3
End If
End Sub


Private Sub chk5frames_Click()
Dim Path As String
    Path = App.Path & "\data\options.ini"
    
    Call PutVar(Path, "OPTIONS", "AnimacionAtaque", CStr(chk5frames.Value))
    Options.AnimacionAtaque = chk5frames.Value
    
    SendSpriteAnimAtaqToAll
End Sub

Private Sub chkDropInvItems_Click()
Dim var1 As String
Dim Path As String                                                              '[TIRARINVENTARIO]
    Path = App.Path & "\data\options.ini"
    If chkDropInvItems.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TIRARINVENTARIO", "C1")
        chkDropInvItems.Caption = var1       '[TIRARINVENTARIO]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TIRARINVENTARIO", "C2")
        chkDropInvItems.Caption = var1    '[TIRARINVENTARIO]C2
    End If
    
    Call PutVar(Path, "OPTIONS", "DropOnDeath", CStr(chkDropInvItems.Value))
End Sub

Private Sub chkdropmuerte_Click()
txtprobabilidaddrop.Enabled = Not txtprobabilidaddrop.Enabled
End Sub

Private Sub chkFriendSystem_Click()
Dim Path As String
Dim var1 As String
'[AMIGOS]
    Path = App.Path & "\data\options.ini"
    If chkFriendSystem.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "AMIGOS", "C1")
        chkFriendSystem.Caption = var1   '[AMIGOS]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TIRARINVENTARIO", "C2")
        chkFriendSystem.Caption = var1     '[AMIGOS]C2
    End If
    
    Call PutVar(Path, "OPTIONS", "FriendSystem", CStr(chkFriendSystem.Value))
End Sub

Private Sub chkFS_Click()
Dim Path As String
Dim var1 As String

    Path = App.Path & "\data\options.ini"                   '[PANTALLACOMPLETA]
    If chkFS.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANTALLACOMPLETA", "C1")
        chkFS.Caption = var1            '[PANTALLACOMPLETA]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANTALLACOMPLETA", "C2")
        chkFS.Caption = var1            '[PANTALLACOMPLETA]C2
    End If
    
    Call PutVar(Path, "OPTIONS", "FullScreen", CStr(chkFS.Value))
    Options.FullScreen = chkFS.Value
    
    SendHighIndex
End Sub

Private Sub chkGUIBars_Click()
Dim Path As String
Dim var1 As String

    Path = App.Path & "\data\options.ini"           '[GUIBARS]
    If chkGUIBars.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GUIBARS", "C1")
        chkGUIBars.Caption = var1     '[GUIBARS]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GUIBARS", "C2")
        chkGUIBars.Caption = var1     '[GUIBARS]C2
    End If
    
    Call PutVar(Path, "OPTIONS", "OriginalGUIBars", CStr(chkGUIBars.Value))
    Options.OriginalGUIBars = chkGUIBars.Value
    
    SendGUIBarsToAll
End Sub

Private Sub chkProj_Click()
Dim Path As String
Dim var1 As String

    Path = App.Path & "\data\options.ini"           '[PROYECTILES]
    If chkProj.Value Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PROYECTILES", "C1")
        chkProj.Caption = var1        '[PROYECTILES]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PROYECTILES", "C2")
        chkProj.Caption = var1        '[PROYECTILES]C2
    End If
    
    Call PutVar(Path, "OPTIONS", "Projectiles", CStr(chkProj.Value))
    Options.Projectiles = chkProj.Value
End Sub

Private Sub cmdGuardar_Click()

End Sub

Private Sub chkrandclima_Click()
If chkrandclima.Value = 1 Then
cmbhora0.Enabled = False
cmbhora1.Enabled = False
cmbhora2.Enabled = False
cmbhora3.Enabled = False
cmbhora4.Enabled = False
cmbhora5.Enabled = False
cmbhora6.Enabled = False
cmbhora7.Enabled = False
cmbhora8.Enabled = False
cmbhora9.Enabled = False
cmbhora10.Enabled = False
cmbhora11.Enabled = False
cmbhora12.Enabled = False
cmbhora13.Enabled = False
cmbhora14.Enabled = False
cmbhora15.Enabled = False
cmbhora16.Enabled = False
cmbhora17.Enabled = False
cmbhora18.Enabled = False
cmbhora19.Enabled = False
cmbhora20.Enabled = False
cmbhora21.Enabled = False
cmbhora22.Enabled = False
cmbhora23.Enabled = False

Else

cmbhora0.Enabled = True
cmbhora1.Enabled = True
cmbhora2.Enabled = True
cmbhora3.Enabled = True
cmbhora4.Enabled = True
cmbhora5.Enabled = True
cmbhora6.Enabled = True
cmbhora7.Enabled = True
cmbhora8.Enabled = True
cmbhora9.Enabled = True
cmbhora10.Enabled = True
cmbhora11.Enabled = True
cmbhora12.Enabled = True
cmbhora13.Enabled = True
cmbhora14.Enabled = True
cmbhora15.Enabled = True
cmbhora16.Enabled = True
cmbhora17.Enabled = True
cmbhora18.Enabled = True
cmbhora19.Enabled = True
cmbhora20.Enabled = True
cmbhora21.Enabled = True
cmbhora22.Enabled = True
cmbhora23.Enabled = True

End If
End Sub

Private Sub CmdAtrasClima_Click()
CmdAtrasClima.Picture = LoadPicture(App.Path & "\data\GUI\Atras_S.Jpg")
Pricmaquina.Picture = LoadPicture(App.Path & "\data\GUI\Maquina.Jpg")
PicClima.Visible = True
PicMaquinaClima.Visible = False
End Sub

Private Sub CmdAtrasClima_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
                            '[CLIMA]
                            Dim var1 As String
                            
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C5")
LblEstado.Caption = var1 '[CLIMA]C5
End Sub

Private Sub Cmdboton_Click(index As Integer)
Dim ImgbotonS As String
Dim ImgBoton As String

If tmrGetTime.Enabled = False Then Exit Sub

PicJugadores.Visible = False
Picconsola.Visible = False
PicControl.Visible = False
PicEditor.Visible = False

Cmdboton(0).Picture = LoadPicture(App.Path & "\data\GUI\0.Jpg")
Cmdboton(1).Picture = LoadPicture(App.Path & "\data\GUI\1.Jpg")
Cmdboton(2).Picture = LoadPicture(App.Path & "\data\GUI\2.Jpg")
Cmdboton(3).Picture = LoadPicture(App.Path & "\data\GUI\3.Jpg")

Select Case index
 Case 0  'boton consola
               'Animacion de boton
               Cmdboton(index).Picture = LoadPicture(App.Path & "\data\GUI\" & index & "_S.Jpg")
               Cmdboton(1).Picture = LoadPicture(App.Path & "\data\GUI\1.Jpg")
               Cmdboton(2).Picture = LoadPicture(App.Path & "\data\GUI\2.Jpg")
               Cmdboton(3).Picture = LoadPicture(App.Path & "\data\GUI\3.Jpg")
               
               'activando ventanas
               Picconsola.Visible = True
               txtText.Visible = True
      
      Case 1 'boton jugador
               
               'Animacion de boton
               Cmdboton(index).Picture = LoadPicture(App.Path & "\data\GUI\" & index & "_S.Jpg")
      
               'activando ventanas
               PicJugadores.Visible = True
               txtText.Visible = False
      Case 2 'boton Control
               
               'Animacion de boton
               Cmdboton(index).Picture = LoadPicture(App.Path & "\data\GUI\" & index & "_S.Jpg")
      
               'activando ventanas
               PicControl.Visible = True
               txtText.Visible = False
      Case 3 'Boton Editor
               
               'Animacion de boton
               Cmdboton(index).Picture = LoadPicture(App.Path & "\data\GUI\" & index & "_S.Jpg")
      
               'activando ventanas
               PicEditor.Visible = True
               txtText.Visible = False
      
End Select

txtText.Top = 1560
txtText.Left = 3840
txtText.Height = 3615
txtText.Width = 7155

PicJugadores.Top = 1320
PicControl.Top = 1320
PicEditor.Top = 1320

PicJugadores.Left = 3720
PicControl.Left = 3720
PicEditor.Left = 3720

End Sub

Private Sub Cmdboton_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

If tmrGetTime.Enabled = False Then Exit Sub
Select Case index
Case 0                                      '[PANELCONTROL]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANELCONTROL", "C1")
LblEstado.Caption = var1       '[PANELCONTROL]C1
Case 1
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANELCONTROL", "C2")
LblEstado.Caption = var1  '[PANELCONTROL]C2
Case 2
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANELCONTROL", "C3")
LblEstado.Caption = var1   '[PANELCONTROL]C3
Case 3
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANELCONTROL", "C4")
LblEstado.Caption = var1              '[PANELCONTROL]C4
End Select
End Sub

Private Sub cmdclima_Click()
Dim Clima As Long
Dim intensidad As Long
Dim mapa As Long
Dim niebla As Long
Dim Velocidad As Long
Dim opacidad As Long
Dim var1 As String

cmdclima.Picture = LoadPicture(App.Path & "\data\GUI\Generar_S.Jpg")

If Len(txtmapaclima.text) < 1 Or Len(txtintensidadclima.text) < 1 Or cmbclima.ListIndex < 0 Or Len(txtniebla.text) < 1 Or Len(txtvelocidad.text) < 1 Or Len(txtopacidad.text) < 1 Then
                                                                                        '[CLIMA]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "M1")
'MsgBox (var1)                                             '[CLIMA]M1
Else
Clima = cmbclima.ListIndex
intensidad = txtintensidadclima.text
mapa = txtmapaclima.text
niebla = txtniebla.text
opacidad = txtopacidad.text
Velocidad = txtvelocidad.text
If intensidad > 100 Or niebla > 255 Or opacidad > 255 Or Velocidad > 255 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "M2")
'MsgBox (var1)  '[CLIMA]M2
Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "G1")
Call GlobalMsg(var1, Red)                                       '[CLIMA]G1
Call SendClima(mapa, Clima, intensidad, niebla, Velocidad, opacidad)
End If
End If
cmdclima.Picture = LoadPicture(App.Path & "\data\GUI\Generar.Jpg")
End Sub

Private Sub cmdclima_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

                                        '[CLIMA]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLIMA", "C6")
LblEstado.Caption = var1     '[CLIMA]C6
End Sub

Private Sub CmdEdit_Click(index As Integer)
'ocultar menues
PicMaximos.Visible = False
PicClima.Visible = False
PicMenumuerte.Visible = False
PicGremio.Visible = False
PicEditClases.Visible = False
PicMaquinaClima.Visible = False

'imagenes normales
CmdEdit(0).Picture = LoadPicture(App.Path & "\data\GUI\4.Jpg")
CmdEdit(1).Picture = LoadPicture(App.Path & "\data\GUI\17.Jpg")
CmdEdit(2).Picture = LoadPicture(App.Path & "\data\GUI\16.Jpg")
CmdEdit(3).Picture = LoadPicture(App.Path & "\data\GUI\18.Jpg")
CmdEdit(4).Picture = LoadPicture(App.Path & "\data\GUI\5.Jpg")

Select Case index
       Case 0 'Clima
          PicClima.Visible = True
          CmdAtrasClima.Picture = LoadPicture(App.Path & "\data\GUI\Atras.Jpg")
          CmdEdit(0).Picture = LoadPicture(App.Path & "\data\GUI\4_S.Jpg")
       Case 1 'Morir
              If chkbloqpj.Value = 0 Then
              ChckMenuMuerte(0).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
              Else
              ChckMenuMuerte(0).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
              End If
              If chkdropmuerte.Value = 0 Then
              ChckMenuMuerte(1).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
              Else
              ChckMenuMuerte(1).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
              End If
              PicMenumuerte.Visible = True
              CmdEdit(1).Picture = LoadPicture(App.Path & "\data\GUI\17_S.Jpg")
              
          
       Case 2 'Maximos
          PicMaximos.Visible = True
          CmdEdit(2).Picture = LoadPicture(App.Path & "\data\GUI\16_S.Jpg")
       Case 3 'Gremios
          PicGremio.Visible = True
          CmdEdit(3).Picture = LoadPicture(App.Path & "\data\GUI\18_S.Jpg")
       Case 4
          PicEditClases.Visible = True
          CmdEdit(4).Picture = LoadPicture(App.Path & "\data\GUI\5_S.Jpg")
          Dim filename As String
          filename = App.Path & "\data\classes.ini"
          If chkvisible.Value = 1 Then
          CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
          Else
          CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
End If
End Select
End Sub

Private Sub CmdEdit_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

Select Case index
Case 0                                                      '[EDITORES]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EDITORES", "C1")
LblEstado.Caption = var1   '[EDITORES]C1
Case 1
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EDITORES", "C2")
LblEstado.Caption = var1                     '[EDITORES]C2
Case 2
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EDITORES", "C3")
LblEstado.Caption = var1                               '[EDITORES]C3
Case 3
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EDITORES", "C4")
LblEstado.Caption = var1   '[EDITORES]C4
Case 4
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "EDITORES", "C5")
LblEstado.Caption = var1                         '[EDITORES]C5
End Select
End Sub

Private Sub cmdGSave_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "OTROS", "C4")
LblEstado.Caption = var1 '[OTROS]C4
End Sub

Private Sub cmdguardarclases_Click()
Dim var1 As String

cmdguardarclases.Picture = LoadPicture(App.Path & "\data\GUI\Ok_S.Jpg")
If txtnombreclase.text = "" Or txtspritemascclase.text = "" Or txtspritefemclase.text = "" Or txtfuerzaclase.text = "" Or txtresistenciaclase.text = "" Or txtinteligenciaclase.text = "" Or txtagilidadclase.text = "" Or txtvoluntadclase.text = "" Or txtmapaspawn.text = "" Or txtxspawn.text = "" Or txtyspawn.text = "" Or txtcaminarvelocidad = "" Or txtcorrervelocidad = "" Then

                                            '[CLASES]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "M1")
'MsgBox (var1)    '[CLASES]M1
Else

PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Name", txtnombreclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "MaleSprite", txtspritemascclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "FemaleSprite", txtspritefemclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Strength", txtfuerzaclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Endurance", txtresistenciaclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Intelligence", txtinteligenciaclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Agility", txtagilidadclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "WillPower", txtvoluntadclase.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Mapa", txtmapaspawn.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "X", txtxspawn.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Y", txtyspawn.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "VCaminar", txtcaminarvelocidad.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "VCorrer", txtcorrervelocidad.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Visible", chkvisible.Value
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartItem1", txtItemN1.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartItem2", txtItemN2.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartItem3", txtItemN3.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartValue1", txtItemN1.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartValue2", txtItemN2.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartValue3", txtItemN3.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartSpell1", txtSpellN1.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartSpell2", txtSpellN2.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "StartSpell3", txtSpellN3.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "Faccion", hscrollfaccion.Value
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "ItemFaccion", txtitemfaccion.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "NivEvol", txtevolucionniv.text
PutVar App.Path & "\data\classes.ini", "CLASS" & sldclasenum.Value, "ClaseEvol", txtevolclase.text

End If
cmdguardarclases.Picture = LoadPicture(App.Path & "\data\GUI\Ok.Jpg")
End Sub







Private Sub cmdguardarclases_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
                                                '[CLASES]
                                                Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "C1")
LblEstado.Caption = var1   '[CLASES]C1
End Sub
                                                '[RECARGAR]
Private Sub cmdReloadAnimations_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C1")
LblEstado.Caption = var1      '[RECARGAR]C1
End Sub

Private Sub cmdReloadClasses_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C2")
LblEstado.Caption = var1           '[RECARGAR]C2
End Sub

Private Sub cmdReloadCombos_Click()
Dim I As Long
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T1")
   cmdReloadCombos.Picture = LoadPicture(App.Path & "\data\GUI\13_S.Jpg")
    Call LoadCombos
    Call TextAdd(var1)        '[RECARGAR]T1
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendCombos I
        End If
    Next
   cmdReloadCombos.Picture = LoadPicture(App.Path & "\data\GUI\13.Jpg")
End Sub

Private Sub cmdReloadCombos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C3")
LblEstado.Caption = var1          '[RECARGAR]C3
End Sub

Private Sub cmdReloadItems_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C4")
LblEstado.Caption = var1          '[RECARGAR]C4
End Sub

Private Sub cmdReloadMaps_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C5")
LblEstado.Caption = var1            '[RECARGAR]C5
End Sub

Private Sub cmdReloadNPCs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C6")
LblEstado.Caption = var1              '[RECARGAR]C6
End Sub

Private Sub cmdReloadQuests_Click()
Dim I As Long
    cmdReloadQuests.Picture = LoadPicture(App.Path & "\data\GUI\7_S.Jpg")
    Call LoadQuests
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T2")
    Call TextAdd(var1)      '[RECARGAR]T2
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendQuests I
        End If
    Next
   cmdReloadQuests.Picture = LoadPicture(App.Path & "\data\GUI\7.Jpg")
End Sub
Private Sub cmdGSave_Click()
    cmdGSave.Picture = LoadPicture(App.Path & "\data\GUI\Ok_S.Jpg")
    Options.Buy_Cost = frmServer.txtGBuyCost.text
    Options.Buy_Lvl = frmServer.txtGBuyLvl.text
    Options.Buy_Item = frmServer.txtGBuyItem.text
    Options.Join_Cost = frmServer.txtGJoinCost.text
    Options.Join_Lvl = frmServer.txtGJoinLvl.text
    Options.Join_Item = frmServer.txtGJoinItem.text
    SaveOptions
    cmdGSave.Picture = LoadPicture(App.Path & "\data\GUI\Ok.Jpg")
End Sub

Private Sub cmdtest_Click()
Time_Minutes = 59
Time_Seconds = 55

End Sub

Private Sub cmdReloadQuests_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C7")
LblEstado.Caption = var1     '[RECARGAR]C7
End Sub

Private Sub cmdReloadResources_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C8")
LblEstado.Caption = var1     '[RECARGAR]C8
End Sub

Private Sub cmdReloadShops_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C9")
LblEstado.Caption = var1      '[RECARGAR]C9
End Sub

Private Sub CmdReloadSpells_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "C10")
LblEstado.Caption = var1     '[RECARGAR]C10
End Sub

Private Sub cmdShutDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
If tmrGetTime.Enabled = False Then Exit Sub
If isShuttingDown Then                                          '[APAGAR]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "APAGAR", "C1")
        LblEstado.Caption = var1     '[APAGAR]C1
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "APAGAR", "C2")
        LblEstado.Caption = var1                   '[APAGAR]C2
    End If
End Sub



Private Sub hscrollfaccion_Change()
LblFacionNumero.Caption = hscrollfaccion.Value
End Sub

Private Sub ImgLogo_Click()
If tmrGetTime.Enabled = False Then Exit Sub
ShellExecute hWnd, "open", "http://easee.es/", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub ImgLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If tmrGetTime.Enabled = False Then Exit Sub
LblEstado.Caption = "WWW.EASEE.ES"
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub lblCPSLock_Click()
Dim var1 As String
If tmrGetTime.Enabled = False Then Exit Sub
    If CPSUnlock Then
        CPSUnlock = False                   '[CPS]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CPS", "C1")
        lblCpsLock.Caption = var1   '[CPS]C1
    Else
        CPSUnlock = True
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CPS", "C2")
        lblCpsLock.Caption = var1    '[CPS]C2
    End If
End Sub


Private Sub lvwInfo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANELCONTROL", "C5")
LblEstado.Caption = var1 '[PANELCONTROL]C5
End Sub

Private Sub PicActualizar_Click()

End Sub

Private Sub PicControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
LblEstado.Caption = ""
End Sub

Private Sub Pickcheck_Click(index As Integer)
Dim Path As String
Path = App.Path & "\data\options.ini"

Select Case index

Case 0 ' Sistema de amigos
         If chkFriendSystem.Value = 1 Then
         chkFriendSystem.Value = 0
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Else
         chkFriendSystem.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         End If
         
Case 1 ' Proyectiles
         If chkProj.Value = 1 Then
         chkProj.Value = 0
          Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Else
         chkProj.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         End If
         
Case 2 ' Pantalla completa

         
Case 3 ' Inventario al morir
         If chkDropInvItems.Value = 1 Then
         chkDropInvItems.Value = 0
          Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Else
         chkDropInvItems.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         End If
         
Case 4 ' GUI
         If chkGUIBars.Value = 1 Then
         chkGUIBars.Value = 0
          Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Else
         chkGUIBars.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         End If
                  
Case 5 ' LOGS
         If chkServerLog.Value = 1 Then
         chkServerLog.Value = 0
          Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Else
         chkServerLog.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         End If
Case 6 ' sprites
         If chk5frames.Value = 1 Then
         chk5frames.Value = 0
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         Options.AnimacionAtaque = 0
         Call PutVar(Path, "OPTIONS", "AnimacionAtaque", 0)
         Else
         Call PutVar(Path, "OPTIONS", "AnimacionAtaque", 1)
         chk5frames.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         Options.AnimacionAtaque = 1
         End If
         SendSpriteAnimAtaqToAll
        
Case 7 ' Ataque aliado version 0.9 EaSee
         If chkataliado.Value = 1 Then
         chkataliado.Value = 0
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
         'Options.AnimacionAtaque = 0
         Call PutVar(Path, "OPTIONS", "AtaqueAliado", 0)
         Else
         Call PutVar(Path, "OPTIONS", "AtaqueAliado", 1)
         chkataliado.Value = 1
         Pickcheck(index).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         'Options.AnimacionAtaque = 1
         End If
        

         
         End Select
         
End Sub

Private Sub Pickcheck_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

Select Case index

Case 0 ' Sistema de amigos
         If chkFriendSystem.Value Then              '[AMIGOS]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "AMIGOS", "C3")
         LblEstado.Caption = var1   '[AMIGOS]C3
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "AMIGOS", "C4")
         LblEstado.Caption = var1      '[AMIGOS]C4
         End If
         
Case 1 ' Proyectiles
         If chkProj.Value Then                      '[PROYECTILES]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PROYECTILES", "C3")
         LblEstado.Caption = var1   '[PROYECTILES]C3
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PROYECTILES", "C4")
         LblEstado.Caption = var1      '[PROYECTILES]C4
         End If
         
Case 2 ' Pantalla completa
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "PANTALLACOMPLETA", "C3")
         LblEstado.Caption = var1 '[PANTALLACOMPLETA]C3
         
Case 3 ' Inventario al morir
         If chkDropInvItems.Value Then              '[TIRARINVENTARIO]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TIRARINVENTARIO", "C3")
         LblEstado.Caption = var1   '[TIRARINVENTARIO]C3
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "TIRARINVENTARIO", "C4")
         LblEstado.Caption = var1      '[TIRARINVENTARIO]C4
         End If
         
Case 4 ' GUI
         If chkGUIBars.Value Then                                   '[GUIBARS]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GUIBARS", "C3")
         LblEstado.Caption = var1           '[GUIBARS]C3
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "GUIBARS", "C4")
         LblEstado.Caption = var1    '[GUIBARS]C4
         End If
                  
Case 5 ' LOGS
         If chkServerLog.Value Then                                 '[LOGS]
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "C1")
         LblEstado.Caption = var1                           '[LOGS]C1
         Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "LOGS", "C2")
         LblEstado.Caption = var1                              '[LOGS]C2
         End If
         
         End Select
End Sub



Private Sub Pricmaquina_Click()
Pricmaquina.Picture = LoadPicture(App.Path & "\data\GUI\Maquina_S.Jpg")
CmdAtrasClima.Picture = LoadPicture(App.Path & "\data\GUI\Atras.Jpg")
PicClima.Visible = False
PicMaquinaClima.Visible = True



If chkrandclima.Value = 0 Then
      checkclima(0).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
    Else
      checkclima(0).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
    End If
    If chkclimas.Value = 0 Then
      checkclima(1).Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
    Else
      checkclima(1).Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
      End If
      
Pricmaquina.Picture = LoadPicture(App.Path & "\data\GUI\Maquina.Jpg")
End Sub

Private Sub Pricmaquina_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "OTROS", "c1")
LblEstado.Caption = var1    '[OTROS]C1
End Sub

Private Sub sldclasenum_Change() 'Cortesia de EaSee Engine (que lindo)
Dim filename As String
filename = App.Path & "\data\classes.ini"
lblnumero.Caption = (sldclasenum.Value)
frmServer.txtnombreclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Name")
        frmServer.txtspritemascclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "MaleSprite")
        frmServer.txtspritefemclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "FemaleSprite")
        frmServer.txtfuerzaclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Strength")
        frmServer.txtresistenciaclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Endurance")
        frmServer.txtinteligenciaclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Intelligence")
        frmServer.txtagilidadclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Agility")
        frmServer.txtvoluntadclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "WillPower")
        frmServer.txtmapaspawn.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Mapa")
        frmServer.txtxspawn.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "X")
        frmServer.txtyspawn.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Y")
        frmServer.txtcaminarvelocidad.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "VCaminar")
        frmServer.txtcorrervelocidad.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "VCorrer")
        frmServer.chkvisible.Value = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Visible")
        frmServer.txtItemN1.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartItem1")
        frmServer.txtItemN2.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartItem2")
        frmServer.txtItemN3.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartItem3")
        frmServer.txtItemC1.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartValue1")
        frmServer.txtItemC2.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartValue2")
        frmServer.txtItemC3.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartValue3")
        frmServer.txtSpellN1.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartSpell1")
        frmServer.txtSpellN2.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartSpell2")
        frmServer.txtSpellN3.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "StartSpell3")
        frmServer.hscrollfaccion.Value = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Faccion") 'EaSee 0.6
        frmServer.txtevolucionniv.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "NivEvol") 'EaSee 0.7
        frmServer.txtevolclase.text = GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "ClaseEvol")

        
        
If GetVar(filename, "CLASS" & frmServer.sldclasenum.Value, "Visible") = 1 Then
         CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\ON.Jpg")
         Else
         CheckVisibleClase.Picture = LoadPicture(App.Path & "\data\GUI\OFF.Jpg")
End If
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
Dim Path As String
Path = App.Path & "\data\options.ini"
    ' if its not 0, then its true
    If Not chkServerLog.Value Then
        ServerLog = True
    End If
Call PutVar(Path, "OPTIONS", "Logs", CStr(chkServerLog.Value))
End Sub

Private Sub cmdExit_Click()
    Call DestroyServer
End Sub

Private Sub cmdReloadClasses_Click()
Dim I As Long
Dim var1 As String
    cmdReloadClasses.Picture = LoadPicture(App.Path & "\data\GUI\5_S.Jpg")
    Call LoadClasses
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "T1")
    Call TextAdd(var1)    '[CLASES]T1
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendClasses I
            
        End If
    Next
    cmdReloadClasses.Picture = LoadPicture(App.Path & "\data\GUI\5.Jpg")
    
End Sub

Private Sub cmdReloadItems_Click()
Dim I As Long
Dim var1 As String
    cmdReloadItems.Picture = LoadPicture(App.Path & "\data\GUI\6_S.Jpg")
    Call LoadItems
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T3")
    Call TextAdd(var1)       '[RECARGAR]T3
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendItems I
        End If
    Next
    cmdReloadItems.Picture = LoadPicture(App.Path & "\data\GUI\6.Jpg")
End Sub

Private Sub cmdReloadMaps_Click()
Dim I As Long
Dim var1 As String
 cmdReloadMaps.Picture = LoadPicture(App.Path & "\data\GUI\11_S.Jpg")
    Call LoadMaps
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T4")
    Call TextAdd(var1)       '[RECARGAR]T4
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            PlayerWarp I, GetPlayerMap(I), GetPlayerX(I), GetPlayerY(I)
        End If
    Next
    cmdReloadMaps.Picture = LoadPicture(App.Path & "\data\GUI\11.Jpg")
End Sub

Private Sub cmdReloadNPCs_Click()
Dim I As Long
Dim var1 As String
    cmdReloadNPCs.Picture = LoadPicture(App.Path & "\data\GUI\9_S.Jpg")
    Call LoadNpcs
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T5")
    Call TextAdd(var1)       '[RECARGAR]T5
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendNpcs I
        End If
    Next
    cmdReloadNPCs.Picture = LoadPicture(App.Path & "\data\GUI\9.Jpg")
End Sub

Private Sub cmdReloadShops_Click()
Dim I As Long
Dim var1 As String

   cmdReloadShops.Picture = LoadPicture(App.Path & "\data\GUI\12_S.Jpg")
    Call LoadShops
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T6")
    Call TextAdd(var1)       '[RECARGAR]T6
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendShops I
        End If
    Next
    cmdReloadShops.Picture = LoadPicture(App.Path & "\data\GUI\12.Jpg")
End Sub

Private Sub cmdReloadSpells_Click()
Dim I As Long
Dim var1 As String
    CmdReloadSpells.Picture = LoadPicture(App.Path & "\data\GUI\10_S.Jpg")
    Call LoadSpells
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T7")
    Call TextAdd(var1)       '[RECARGAR]T7
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendSpells I
        End If
    Next
   CmdReloadSpells.Picture = LoadPicture(App.Path & "\data\GUI\10.Jpg")
End Sub

Private Sub cmdReloadResources_Click()
Dim I As Long
Dim var1 As String

    cmdReloadResources.Picture = LoadPicture(App.Path & "\data\GUI\8_S.Jpg")
    Call LoadResources
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T8")
    Call TextAdd(var1)       '[RECARGAR]T8
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendResources I
        End If
    Next
    cmdReloadResources.Picture = LoadPicture(App.Path & "\data\GUI\8.Jpg")
End Sub

Private Sub cmdReloadAnimations_Click()
Dim I As Long
Dim var1 As String

    cmdReloadAnimations.Picture = LoadPicture(App.Path & "\data\GUI\14_S.Jpg")
    Call LoadAnimations
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "RECARGAR", "T9")
    Call TextAdd(var1)       '[RECARGAR]T9
    For I = 1 To Player_HighIndex
        If isPlaying(I) Then
            SendAnimations I
        End If
    Next
    cmdReloadAnimations.Picture = LoadPicture(App.Path & "\data\GUI\14.Jpg")
End Sub

Private Sub cmdShutDown_Click()
Dim var1 As String

If tmrGetTime.Enabled = False Then Exit Sub
    If isShuttingDown Then
        isShuttingDown = False
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "APAGAR", "G1")
        GlobalMsg var1, BrightBlue          '[APAGAR]G1
        frmServer.LblApagagando.Visible = False
        cmdShutDown.Picture = LoadPicture(App.Path & "\data\GUI\Apagar.Jpg")
    Else
        isShuttingDown = True
        frmServer.LblApagagando.Visible = True
        cmdShutDown.Picture = LoadPicture(App.Path & "\data\GUI\Apagado_S.Jpg")
    End If
End Sub

Private Sub Form_Load()
Dim Extra As String
    Cmdboton(0).Picture = LoadPicture(App.Path & "\data\GUI\0_S.Jpg")
    Call PreLoad
    Extra = command$
    If Extra = "-#Mode$Singer$#" Then
    If Me.Visible = True Then Me.Visible = False
    Timer1.Enabled = True
    txtText.Height = Me.Height - 400
    txtText.Width = Me.Width - 100
    txtText.Left = 0
    txtText.Top = 0
    Command1.Visible = True
    Me.WindowState = 0
    Else
    If GetVar(App.Path & "\data\options.ini", "SINGLE", "Activado") = "1" Then
    MsgBox "Error en el motor SinglePlayer", vbCritical, "Server error"
    End
    End If
    End If
    
    MDE = "0"
    
    Call SetData
    
    'Cargar idioma
    Language = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Lang")
    
    LblMenumorir(0).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c1")
    LblMenumorir(1).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c2")
    LblMenumorir(2).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c3")
    Label15.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c4")
    Label6.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c5")
    Label1.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c6")
    Label2.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c7")
    Label16.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c8")
    Label5.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c9")
    Label3.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c10")
    Label4.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c11")
    Label13.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c12")
    lblmapa.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c13")
    lbltipoclima.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c14")
    Label7.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c15")
    Label8.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c16")
    Label10.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c17")
    Label12.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c18")
    Lblcheckclima(1).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c19")
    Lblcheckclima(0).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c20")
    Label25.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c21")
    Label26.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c22")
    lblnombre.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c23")
    lblStats(0).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c24")
    lblStats(3).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c25")
    lblStats(4).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c26")
    lblStats(1).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c27")
    lblStats(2).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c28")
    lvlvelocidad.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c29")
    lblcaminar.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c30")
    lblcorrer.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c31")
    Label23.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c32")
    Label24.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c33")
    Label17.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c34")
    Label22.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c35")
    Label18.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c36")
    Label19.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c37")
    Label20.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c38")
    lblsprite(0).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c39")
    lblsprite(1).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c40")
    lblspawn.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c41")
    lblmapaspawn.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c42")
    LblCheck(0).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c43")
    LblCheck(1).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c44")
    LblCheck(3).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c45")
    LblCheck(4).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c46")
    LblCheck(2).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c47")
    LblCheck(6).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c48")
    LblCheck(5).Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c49")
    Label11.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c50")
    lblCpsLock.Caption = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c51")
    Cmdboton(0).ToolTipText = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c52")
    Cmdboton(1).ToolTipText = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c53")
    Cmdboton(2).ToolTipText = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c54")
    Cmdboton(3).ToolTipText = GetVar(App.Path & "\data\lang\" & Language & ".ini", "FORM", "c55")
    
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

Private Sub Timer1_Timer()
    If Me.Visible = True Then Me.Visible = False
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
            Time_Seconds = 0
        If frmServer.chkclimas.Value = 0 Then
            Call Procesar_Clima
        End If
        If Time_Hours + 1 = 24 Then
        Time_Hours = 0
        End If
        Else
            If MDE = 1 Then
            If frmServer.Socket(1).State = 7 Then Else Call DestroyServer
            End If
            Time_Minutes = Time_Minutes + 1
            Time_Seconds = 0
        End If
    Else
        Time_Seconds = Time_Seconds + 1
        If MDE = 1 Then
        'If Time_Hours > 0 Or Time_Minutes > 0 Then
         If frmServer.Socket(1).State = 7 Then Else Call DestroyServer
        'If DebugSingle = 0 Then
        'End If
        End If
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

Private Sub txtevolclase_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtevolucionniv_KeyPress(KeyAscii As Integer)
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
Dim var1 As String

txtguardarmaxmin.Picture = LoadPicture(App.Path & "\data\GUI\OK_S.Jpg")

filename = App.Path & "\data\classes.ini"


If Len(txtmaxclases.text) < 1 Then 'Guardado de Max_Classes
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "M1")
'MsgBox (var1)                     '[CLASES]M1
Else
clasesdeseadas = txtmaxclases.text
PutVar App.Path & "\data\classes.ini", "INIT", "MaxClasses", txtmaxclases.text

If clasesexistentes < clasesdeseadas Then 'Si pusiste mas Clases lo detecta y crea datos del fichero
  
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
        Call PutVar(filename, "CLASS" & I, "Faccion", "1")
        Call PutVar(filename, "CLASS" & I, "ItemFaccion", "0")
        Call PutVar(filename, "CLASS" & I, "NivEvol", "0")
        Call PutVar(filename, "CLASS" & I, "ClaseEvol", "1")

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
    Max_Classes = txtmaxclases.text
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "M2")
    'MsgBox (var1)             '[CLASES]M2
    Else
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "M2")
    'MsgBox (var1)             '[CLASES]M2
End If
End If
txtguardarmaxmin.Picture = LoadPicture(App.Path & "\data\GUI\OK.Jpg")
End Sub



Private Sub txtguardarmaxmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String

var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "CLASES", "C2")
LblEstado.Caption = var1       '[CLASES]C2
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


Private Sub txtitemfaccion_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtItemN1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtItemN2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtItemN3_KeyPress(KeyAscii As Integer)
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



Private Sub txtmaxclases_KeyPress(KeyAscii As Integer)
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

Private Sub txtprobabilidaddrop_KeyPress(KeyAscii As Integer)
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
Private Sub txtItemC1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtItemC2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtItemC3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpellN1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpellN2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtSpellN3_KeyPress(KeyAscii As Integer)
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
        If LenB(Trim$(txtChat.text)) > 0 Then
            Call GlobalMsg(txtChat.text, White)
            Call TextAdd("Server: " & txtChat.text)
            txtChat.text = vbNullString
        End If

        KeyAscii = 0
    End If

End Sub

Sub UsersOnline_Start()
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        frmServer.lvwInfo.ListItems.Add (I)

        If I < 10 Then
            frmServer.lvwInfo.ListItems(I).text = "00" & I
        ElseIf I < 100 Then
            frmServer.lvwInfo.ListItems(I).text = "0" & I
        Else
            frmServer.lvwInfo.ListItems(I).text = I
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
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "A4")
        Call AlertMsg(FindPlayer(Name), var1) '[KICKBANADMIN]A4
    End If

End Sub

Sub mnuDisconnectPlayer_Click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        CloseSocket (FindPlayer(Name))
    End If

End Sub

Sub mnuBanPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call ServerBanIndex(FindPlayer(Name))
    End If

End Sub

Sub mnuAdminPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 4)
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "M1")
        Call SendPlayerData(FindPlayer(Name))           '[KICKBANADMIN]M1
        Call PlayerMsg(FindPlayer(Name), var1, BrightCyan)
    End If

End Sub

Sub mnuRemoveAdmin_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 0)
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "M2")
        Call SendPlayerData(FindPlayer(Name))           '[KICKBANADMIN]M2
        Call PlayerMsg(FindPlayer(Name), var1, BrightRed)
    End If

End Sub

Sub mnuModPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 1)
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "M3")
        Call SendPlayerData(FindPlayer(Name))           '[KICKBANADMIN]M3
        Call PlayerMsg(FindPlayer(Name), var1, BrightCyan)
    End If

End Sub

Sub mnuMapPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 2)
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "M4")
        Call SendPlayerData(FindPlayer(Name))           '[KICKBANADMIN]M4
        Call PlayerMsg(FindPlayer(Name), var1, BrightCyan)
    End If

End Sub

Sub mnuDevPlayer_click()
    Dim Name As String
    Name = frmServer.lvwInfo.SelectedItem.SubItems(3)
Dim var1 As String

    If Not Name = "Fuera de FLeer(FNum)" And Not FindPlayer(Name) = 0 Then
        Call SetPlayerAccess(FindPlayer(Name), 3)
        Call SendPlayerData(FindPlayer(Name))               '[KICKBANADMIN]M5
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "KICKBANADMIN", "M5")
        Call PlayerMsg(FindPlayer(Name), var1, BrightCyan)
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lmsg As Long
    lmsg = x / Screen.TwipsPerPixelX

    Select Case lmsg
        Case WM_LBUTTONDBLCLK
            frmServer.WindowState = vbNormal
            frmServer.Show
            txtText.SelStart = Len(txtText.text)
    End Select
    
    If tmrGetTime.Enabled = False Then Exit Sub
    LblEstado.Caption = "Easee Engine"

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

Private Sub updatebtn_Click()
'EaSee actualizador desactivado
'Dim prevdir
'    prevdir = Mid$(App.Path, 1, InStrRev(App.Path, "\") - 1)
'    prevdir = prevdir & "\EaSeeUpdater.exe"
'    Shell (prevdir), vbNormalFocus
'    Unload Me
    ShellExecute hWnd, "open", "http://easee.es/index.php?board=127.0", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub updatebtn_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim var1 As String
var1 = GetVar(App.Path & "\data\lang\" & Language & ".ini", "OTROS", "C3")
LblEstado.Caption = var1        '[OTROS]C3
End Sub


