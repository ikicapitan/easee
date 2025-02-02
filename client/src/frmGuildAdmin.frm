VERSION 5.00
Begin VB.Form frmGuildAdmin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clan Control"
   ClientHeight    =   5175
   ClientLeft      =   -15
   ClientTop       =   315
   ClientWidth     =   5340
   Icon            =   "frmGuildAdmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameMainUsers 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Editar Miembros"
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame frameUser 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   4695
         Begin VB.CommandButton cmduser 
            Caption         =   "Guardar"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox txtcomment 
            Height          =   855
            Left            =   840
            TabIndex        =   15
            Top             =   600
            Width           =   3735
         End
         Begin VB.ComboBox cmbRanks 
            Height          =   315
            ItemData        =   "frmGuildAdmin.frx":08CA
            Left            =   840
            List            =   "frmGuildAdmin.frx":08CC
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   230
            Width           =   2295
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario:"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Ranking:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.ListBox listusers 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame frameMainRanks 
      Caption         =   "Edit Ranks"
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame frameranks 
         BorderStyle     =   0  'None
         Caption         =   "Hide Later"
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4815
         Begin VB.OptionButton opAccess 
            Caption         =   "Cannot"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   29
            Top             =   1200
            Width           =   855
         End
         Begin VB.OptionButton opAccess 
            Caption         =   "Can"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   28
            Top             =   840
            Width           =   615
         End
         Begin VB.ListBox listAccess 
            Height          =   1425
            ItemData        =   "frmGuildAdmin.frx":08CE
            Left            =   840
            List            =   "frmGuildAdmin.frx":08D0
            TabIndex        =   26
            Top             =   480
            Width           =   3015
         End
         Begin VB.CommandButton cmdRankSave 
            Caption         =   "Save Rank #10"
            Height          =   255
            Left            =   1680
            TabIndex        =   17
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   840
            TabIndex        =   8
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label4 
            Caption         =   "Access:"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.ListBox listranks 
         Height          =   1425
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones de L�der del Clan"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "Opciones"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Opciones Avanzadas del Clan."
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Usuarios"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         ToolTipText     =   "Listado de Usuarios del Clan."
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ranking"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         ToolTipText     =   "Ranking de los Miembros del Clan."
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frameMainoptions 
      Caption         =   "Edit Options"
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   5055
      Begin VB.TextBox txtGuildColor 
         Height          =   285
         Left            =   1200
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtGuildTag 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtGuildName 
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtMOTD 
         Height          =   975
         Left            =   720
         TabIndex        =   25
         Text            =   "Text3"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.CommandButton cmdoptions 
         Caption         =   "Save Options"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   3840
         Width           =   1215
      End
      Begin VB.HScrollBar scrlRecruits 
         Height          =   255
         Left            =   1800
         Max             =   6
         Min             =   1
         TabIndex        =   21
         Top             =   1920
         Value           =   1
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Guild Tag:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Guild Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Motd:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label lblrecruit 
         Caption         =   "100"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Recruits start at rank:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Guild Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmGuildAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbRanks_Click()
    If listusers.ListIndex > 0 Then
        GuildData.Guild_Members(listusers.ListIndex).Rank = cmbRanks.ListIndex
    End If
End Sub

Private Sub cmdoptions_Click()
    Call GuildSave(1, 1)
End Sub

Private Sub cmdRankSave_Click()
    Call GuildSave(3, listranks.ListIndex)
End Sub

Private Sub cmduser_Click()
     Call GuildSave(2, listusers.ListIndex)
End Sub

Private Sub Command1_Click()
    frameMainRanks.Visible = True
    frameMainUsers.Visible = False
    frameMainoptions.Visible = False
End Sub

Private Sub Command2_Click()
    frameMainRanks.Visible = False
    frameMainUsers.Visible = True
    frameMainoptions.Visible = False
End Sub

Private Sub Command3_Click()
    frameMainRanks.Visible = False
    frameMainUsers.Visible = False
    frameMainoptions.Visible = True
End Sub

Private Sub Form_Load()
 'Load all 3 on load
Call Load_Guild_Admin
 Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String
trad = GetVar(App.Path & Lang, "GuildEditor", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "L1")
Frame1.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "L2")
frameMainUsers.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "L3")
Label2.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "L4")
Label3.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "B1")
Command1.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "B2")
Command2.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "B3")
Command3.Caption = trad
trad = GetVar(App.Path & Lang, "GuildEditor", "B4")
cmduser.Caption = trad

End Sub
Public Sub Load_Guild_Admin()
 Call Load_Menu_Options
 Call Load_Menu_Ranks
 Call Load_Menu_Users
End Sub
Public Sub Load_Menu_Options()
scrlRecruits.max = MAX_GUILD_RANKS
scrlRecruits.Value = GuildData.Guild_RecruitRank
txtGuildColor.text = GuildData.Guild_Color

txtGuildName.text = GuildData.Guild_Name
txtGuildTag.text = GuildData.Guild_Tag
txtMOTD.text = GuildData.Guild_MOTD
End Sub
Public Sub Load_Menu_Ranks()
Dim I As Integer

listranks.Clear
listranks.AddItem ("Select rank to edit...")
For I = 1 To MAX_GUILD_RANKS
    listranks.AddItem ("Rank #" & I & ": " & GuildData.Guild_Ranks(I).name)
Next I

    For I = 0 To 1
        opAccess(I).Visible = False
    Next I

frameranks.Visible = False
listranks.ListIndex = 0


End Sub
Public Sub Load_Menu_Users()
Dim I As Integer

listusers.Clear
listusers.AddItem ("Select user to edit...")

For I = 1 To MAX_GUILD_MEMBERS
    listusers.AddItem ("User #" & I & ": " & GuildData.Guild_Members(I).User_Name)
Next I

cmbRanks.Clear
cmbRanks.AddItem ("Must Select Ranks...")
cmbRanks.ListIndex = 0
For I = 1 To MAX_GUILD_RANKS
    cmbRanks.AddItem (GuildData.Guild_Ranks(I).name)
Next I

frameUser.Visible = False
listusers.ListIndex = 0
End Sub

Private Sub listAccess_Click()
Dim I As Integer

If listAccess.ListIndex = 0 Then
    For I = 0 To 1
        opAccess(I).Visible = False
    Next I
    Exit Sub
Else
    For I = 0 To 1
        opAccess(I).Visible = True
    Next I
End If

    opAccess(GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex)).Value = True
End Sub

Private Sub listranks_Click()
Dim I As Integer
Dim HoldString As String

    If listranks.ListIndex = 0 Then
        frameranks.Visible = False
        Exit Sub
    End If
    
    cmdRankSave.Caption = "Guardar Rank #" & listranks.ListIndex
    txtName.text = GuildData.Guild_Ranks(listranks.ListIndex).name
    
listAccess.Clear
listAccess.AddItem ("Select permission to edit...")
For I = 1 To MAX_GUILD_RANKS_PERMISSION
    If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(I) = 1 Then
        HoldString = "Can"
    Else
        HoldString = "Cannot"
    End If
    listAccess.AddItem (GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(I) & " (" & HoldString & ")")
Next I
    
    For I = 0 To 1
        opAccess(I).Visible = False
    Next I
    
    frameranks.Visible = True
End Sub

Private Sub listusers_Click()
Dim I As Integer
    
    If listusers.ListIndex = 0 Then
        frameUser.Visible = False
        Exit Sub
    End If
    cmduser.Caption = "Guardar usuario #" & listusers.ListIndex
    txtcomment.text = GuildData.Guild_Members(listusers.ListIndex).Comment
    cmbRanks.ListIndex = GuildData.Guild_Members(listusers.ListIndex).Rank

    If Not GuildData.Guild_Members(listusers.ListIndex).User_Name = vbNullString Then
        frameUser.Visible = True
    Else
        frameUser.Visible = False
    End If

End Sub

Private Sub opAccess_Click(Index As Integer)
Dim HoldString As String

 If listranks.ListIndex = 0 Then Exit Sub
 If listAccess.ListIndex = 0 Then Exit Sub
 
 GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = Index
 
    If GuildData.Guild_Ranks(listranks.ListIndex).RankPermission(listAccess.ListIndex) = 1 Then
        HoldString = "Can"
    Else
        HoldString = "Cannot"
    End If
    
    listAccess.List(listAccess.ListIndex) = GuildData.Guild_Ranks(listranks.ListIndex).RankPermissionName(listAccess.ListIndex) & " (" & HoldString & ")"
End Sub

Private Sub scrlRecruits_Change()
    lblrecruit.Caption = scrlRecruits.Value
    GuildData.Guild_RecruitRank = scrlRecruits.Value
End Sub

Private Sub txtcomment_Change()
    If listusers.ListIndex = 0 Then Exit Sub
    
    GuildData.Guild_Members(listusers.ListIndex).Comment = txtcomment.text
End Sub

Private Sub txtGuildColor_Change()
    If txtGuildColor.text = vbNullString Then txtGuildColor.text = 0
    If txtGuildColor.text > 17 Then txtGuildColor.text = 17
    
    GuildData.Guild_Color = txtGuildColor.text
End Sub

Private Sub txtMOTD_Change()
    GuildData.Guild_MOTD = txtMOTD.text
End Sub

Private Sub txtGuildName_Change()
    GuildData.Guild_Name = txtGuildName.text
End Sub

Private Sub txtGuildTag_Change()
    GuildData.Guild_Tag = txtGuildTag.text
End Sub

Private Sub txtName_Change()
If listranks.ListIndex = 0 Then Exit Sub

GuildData.Guild_Ranks(listranks.ListIndex).name = txtName.text
End Sub
