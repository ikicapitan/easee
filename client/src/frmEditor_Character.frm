VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditor_Character 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Personajes"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   5175
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Datos Personaje"
         Height          =   3615
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   4095
         Begin TabDlg.SSTab SSTab1 
            Height          =   3135
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5530
            _Version        =   393216
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            BackColor       =   14737632
            TabCaption(0)   =   "General"
            TabPicture(0)   =   "frmEditor_Character.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Label8"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "Label9"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "Label10"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "Label11"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "Label12"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "Label13"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "Label14"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "Label17"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtELvl"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtEExp"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txtEStr"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "txtEEnd"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "txtEInt"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtEAgi"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).Control(14)=   "txtEWill"
            Tab(0).Control(14).Enabled=   0   'False
            Tab(0).Control(15)=   "txtEPts"
            Tab(0).Control(15).Enabled=   0   'False
            Tab(0).ControlCount=   16
            TabCaption(1)   =   "Skills"
            TabPicture(1)   =   "frmEditor_Character.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdESkillSave"
            Tab(1).Control(1)=   "txtESkillNum"
            Tab(1).Control(2)=   "cmdESkillLoad"
            Tab(1).Control(3)=   "txtESkillExp"
            Tab(1).Control(4)=   "txtESkillLvl"
            Tab(1).Control(5)=   "Label3"
            Tab(1).Control(6)=   "Label2"
            Tab(1).Control(7)=   "Label1"
            Tab(1).ControlCount=   8
            TabCaption(2)   =   "Inventario"
            TabPicture(2)   =   "frmEditor_Character.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "txtEInvNum"
            Tab(2).Control(1)=   "cmdEInvLoad"
            Tab(2).Control(2)=   "txtEItemNum"
            Tab(2).Control(3)=   "txtEItemQty"
            Tab(2).Control(4)=   "cmdEInvSave"
            Tab(2).Control(5)=   "Label15"
            Tab(2).Control(6)=   "Label16"
            Tab(2).Control(7)=   "Label18"
            Tab(2).ControlCount=   8
            TabCaption(3)   =   "Banco"
            TabPicture(3)   =   "frmEditor_Character.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txtEBankNum"
            Tab(3).Control(1)=   "cmdEBankLoad"
            Tab(3).Control(2)=   "txtEBItemNum"
            Tab(3).Control(3)=   "txtEBItemQty"
            Tab(3).Control(4)=   "cmdEBankSave"
            Tab(3).Control(5)=   "Label21"
            Tab(3).Control(6)=   "Label20"
            Tab(3).Control(7)=   "Label19"
            Tab(3).ControlCount=   8
            Begin VB.CommandButton cmdESkillSave 
               Caption         =   "Guardar Skill"
               Height          =   255
               Left            =   -73680
               TabIndex        =   46
               Top             =   2280
               Width           =   1095
            End
            Begin VB.TextBox txtESkillNum 
               Height          =   285
               Left            =   -73320
               TabIndex        =   44
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmdESkillLoad 
               Caption         =   "Chequear Skill"
               Height          =   255
               Left            =   -73680
               TabIndex        =   43
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtESkillExp 
               Height          =   285
               Left            =   -72720
               TabIndex        =   40
               Top             =   1680
               Width           =   1095
            End
            Begin VB.TextBox txtESkillLvl 
               Height          =   285
               Left            =   -74160
               TabIndex        =   39
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtEBankNum 
               Height          =   285
               Left            =   -73920
               TabIndex        =   35
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmdEBankLoad 
               Caption         =   "Comprobar Slot"
               Height          =   255
               Left            =   -74280
               TabIndex        =   34
               Top             =   1200
               Width           =   1455
            End
            Begin VB.TextBox txtEBItemNum 
               Height          =   285
               Left            =   -73800
               TabIndex        =   33
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtEBItemQty 
               Height          =   285
               Left            =   -73800
               TabIndex        =   32
               Top             =   2160
               Width           =   1095
            End
            Begin VB.CommandButton cmdEBankSave 
               Caption         =   "Guardar Slot"
               Height          =   255
               Left            =   -74280
               TabIndex        =   31
               Top             =   2640
               Width           =   1095
            End
            Begin VB.TextBox txtEInvNum 
               Height          =   285
               Left            =   -73920
               TabIndex        =   27
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmdEInvLoad 
               Caption         =   "Comprobar Slot"
               Height          =   255
               Left            =   -74280
               TabIndex        =   26
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtEItemNum 
               Height          =   285
               Left            =   -73800
               TabIndex        =   25
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtEItemQty 
               Height          =   285
               Left            =   -73800
               TabIndex        =   24
               Top             =   2160
               Width           =   1095
            End
            Begin VB.CommandButton cmdEInvSave 
               Caption         =   "Guardar Slot"
               Height          =   255
               Left            =   -74280
               TabIndex        =   23
               Top             =   2640
               Width           =   1095
            End
            Begin VB.TextBox txtEPts 
               Height          =   285
               Left            =   840
               TabIndex        =   14
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtEWill 
               Height          =   285
               Left            =   2400
               TabIndex        =   13
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox txtEAgi 
               Height          =   285
               Left            =   2400
               TabIndex        =   12
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtEInt 
               Height          =   285
               Left            =   840
               TabIndex        =   11
               Top             =   2160
               Width           =   855
            End
            Begin VB.TextBox txtEEnd 
               Height          =   285
               Left            =   2400
               TabIndex        =   10
               Top             =   1200
               Width           =   855
            End
            Begin VB.TextBox txtEStr 
               Height          =   285
               Left            =   840
               TabIndex        =   9
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtEExp 
               Height          =   285
               Left            =   2400
               TabIndex        =   8
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtELvl 
               Height          =   285
               Left            =   960
               TabIndex        =   7
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label3 
               Caption         =   "Skill #:"
               Height          =   255
               Left            =   -74040
               TabIndex        =   45
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label2 
               Caption         =   "Exp:"
               Height          =   255
               Left            =   -73200
               TabIndex        =   42
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "Nivel:"
               Height          =   255
               Left            =   -74760
               TabIndex        =   41
               Top             =   1680
               Width           =   615
            End
            Begin VB.Label Label21 
               Caption         =   "Banco #:"
               Height          =   255
               Left            =   -74760
               TabIndex        =   38
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label20 
               Caption         =   "Objeto #:"
               Height          =   255
               Left            =   -74640
               TabIndex        =   37
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label Label19 
               Caption         =   "Cantidad:"
               Height          =   255
               Left            =   -74640
               TabIndex        =   36
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label15 
               Caption         =   "Inv #:"
               Height          =   255
               Left            =   -74640
               TabIndex        =   30
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label16 
               Caption         =   "Objeto #:"
               Height          =   255
               Left            =   -74640
               TabIndex        =   29
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label Label18 
               Caption         =   "Cantidad:"
               Height          =   255
               Left            =   -74640
               TabIndex        =   28
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label Label17 
               Caption         =   "Pts:"
               Height          =   255
               Left            =   360
               TabIndex        =   22
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label14 
               Caption         =   "Vol:"
               Height          =   255
               Left            =   1920
               TabIndex        =   21
               Top             =   2160
               Width           =   495
            End
            Begin VB.Label Label13 
               Caption         =   "Agi:"
               Height          =   255
               Left            =   1920
               TabIndex        =   20
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label12 
               Caption         =   "Int:"
               Height          =   255
               Left            =   360
               TabIndex        =   19
               Top             =   2160
               Width           =   495
            End
            Begin VB.Label Label11 
               Caption         =   "Res:"
               Height          =   255
               Left            =   1920
               TabIndex        =   18
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label10 
               Caption         =   "Fza:"
               Height          =   255
               Left            =   360
               TabIndex        =   17
               Top             =   1680
               Width           =   495
            End
            Begin VB.Label Label9 
               Caption         =   "Exp:"
               Height          =   255
               Left            =   1920
               TabIndex        =   16
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label8 
               Caption         =   "Nivel:"
               Height          =   255
               Left            =   360
               TabIndex        =   15
               Top             =   720
               Width           =   615
            End
         End
      End
      Begin VB.TextBox txtEName 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdELoad 
         Caption         =   "Cargar"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdESave 
         Caption         =   "Guardar"
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jugador:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmEditor_Character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdELoad_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
        
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 1
    buffer.WriteString txtEName.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdELoad_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdESave_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 2
    buffer.WriteString txtEName.text
    buffer.WriteLong txtELvl.text
    buffer.WriteLong txtEExp.text
    buffer.WriteLong txtEPts.text
    buffer.WriteLong txtEEnd.text
    buffer.WriteLong txtEStr.text
    buffer.WriteLong txtEAgi.text
    buffer.WriteLong txtEInt.text
    buffer.WriteLong txtEWill.text
    buffer.WriteLong txtEInvNum.text
    buffer.WriteLong txtEItemNum.text
    buffer.WriteLong txtEItemQty.text
    buffer.WriteLong txtEBankNum.text
    buffer.WriteLong txtEBItemNum.text
    buffer.WriteLong txtEBItemQty.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdESave_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdESkillLoad_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 3
    buffer.WriteString txtEName.text
    buffer.WriteByte txtESkillNum.text
    buffer.WriteByte txtESkillLvl.text
    buffer.WriteLong txtESkillExp.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEInvLoad_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdESkillSave_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 4
    buffer.WriteString txtEName.text
    buffer.WriteByte txtESkillNum.text
    buffer.WriteByte txtESkillLvl.text
    buffer.WriteLong txtESkillExp.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEInvSave_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEInvLoad_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 5
    buffer.WriteString txtEName.text
    buffer.WriteLong txtEInvNum.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEInvLoad_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEInvSave_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 6
    buffer.WriteString txtEName.text
    buffer.WriteLong txtEInvNum.text
    buffer.WriteLong txtEItemNum.text
    buffer.WriteLong txtEItemQty.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEInvSave_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEBankLoad_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 7
    buffer.WriteString txtEName.text
    buffer.WriteLong txtEBankNum.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEBankLoad_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdEBankSave_Click()
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then Exit Sub
    If txtEName.text = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong CCharEditorCommand
    buffer.WriteByte 8
    buffer.WriteString txtEName.text
    buffer.WriteLong txtEBankNum.text
    buffer.WriteLong txtEBItemNum.text
    buffer.WriteLong txtEBItemQty.text
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdEBankSave_Click", "frmEditor_Character", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CharEditor", "C1")
Me.Caption = trad

trad = GetVar(App.Path & Lang, "CharEditor", "L1")
Label7.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L2")
Frame2.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L3")
Label8.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L4")
Label9.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L5")
Label17.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L6")
Label11.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L7")
Label10.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L8")
Label13.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L9")
Label12.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L10")
Label14.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L11")
Label3.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L12")
Label1.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L13")
Label2.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L14")
SSTab1.TabCaption(2) = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L15")
Label15.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L16")
Label16.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L17")
Label18.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L18")
SSTab1.TabCaption(3) = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L19")
Label21.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L20")
Label20.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "L21")
Label19.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B1")
cmdELoad.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B2")
cmdESave.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B3")
cmdESkillLoad.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B4")
cmdESkillSave.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B5")
cmdEInvLoad.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B6")
cmdEInvSave.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B7")
cmdEBankLoad.Caption = trad
trad = GetVar(App.Path & Lang, "CharEditor", "B8")
cmdEBankSave.Caption = trad

End Sub

