VERSION 5.00
Begin VB.Form frmEditor_Combos 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor de Combinaciones"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12825
   ControlBox      =   0   'False
   Icon            =   "frmEditor_Combos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   367
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recompensas"
      Height          =   2895
      Left            =   3240
      TabIndex        =   18
      ToolTipText     =   "Recompensa o Resultado de la Combinacion"
      Top             =   2520
      Width           =   4695
      Begin VB.HScrollBar scrlSkillExp 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   30
         Top             =   2520
         Width           =   3135
      End
      Begin VB.ComboBox cmbGSkill 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   2160
         Width           =   3135
      End
      Begin VB.HScrollBar scrlIndex 
         Height          =   375
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   21
         Top             =   360
         Value           =   1
         Width           =   4455
      End
      Begin VB.HScrollBar scrlGive 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   20
         Top             =   1080
         Value           =   1
         Width           =   4455
      End
      Begin VB.HScrollBar scrlGiveVal 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   100
         TabIndex        =   19
         Top             =   1680
         Width           =   4455
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label lblSkillExp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Exp: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label lblUseless 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   2160
         Width           =   330
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Convertir Objeto Index: 1"
         Height          =   195
         Left            =   1560
         TabIndex        =   24
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto: "
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Valor: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   9840
      TabIndex        =   17
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   11640
      TabIndex        =   16
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Guardar y Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   5040
      Width           =   4695
   End
   Begin VB.CommandButton cmdSSave 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Requisitos"
      Height          =   4335
      Left            =   8040
      TabIndex        =   7
      ToolTipText     =   "Habilidades y Requerimientos."
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkTake2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tomar este Objeto"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3840
         Width           =   4455
      End
      Begin VB.CheckBox chkTake1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tomar Este Objeto"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2520
         Width           =   4455
      End
      Begin VB.HScrollBar scrlItemVal2 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   37
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox cmbItems2 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3120
         Width           =   3135
      End
      Begin VB.HScrollBar scrlItemVal1 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   33
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ComboBox cmbItems1 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1800
         Width           =   3135
      End
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   13
         Top             =   1320
         Width           =   3135
      End
      Begin VB.HScrollBar scrlSkillLevel 
         Height          =   255
         Left            =   1440
         Max             =   100
         TabIndex        =   11
         Top             =   840
         Width           =   3135
      End
      Begin VB.ComboBox cmbSkill 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblItemVal2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label lblUseless 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto Requerido"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   1245
      End
      Begin VB.Label lblItemVal1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label lblUseless 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Objeto Requerido:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4560
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   540
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4560
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblSLevel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Nivel: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   870
      End
      Begin VB.Label lblUseless 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Propiedades"
      Height          =   2295
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Items/Objetos a Combinar"
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkItem2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tomar Objeto 2"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CheckBox chkItem1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tomar Objeto 1"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   4335
      End
      Begin VB.HScrollBar scrlItem2 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   6
         Top             =   1560
         Value           =   1
         Width           =   4455
      End
      Begin VB.HScrollBar scrlItem1 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   4
         Top             =   600
         Value           =   1
         Width           =   4455
      End
      Begin VB.Label lblSecond 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Segundo Objeto Requerido: "
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2025
      End
      Begin VB.Label lblFirst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Objeto Requerido: "
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lista Objetos"
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   4935
         ItemData        =   "frmEditor_Combos.frx":08CA
         Left            =   120
         List            =   "frmEditor_Combos.frx":08CC
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmEditor_Combos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkItem1_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    Combo(EditorIndex).Take_Item1 = chkItem1.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkItem1_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkItem2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    Combo(EditorIndex).Take_Item2 = chkItem2.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkItem2_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkTake1_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    Combo(EditorIndex).Take_ReqItem1 = chkTake1.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkTake1_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkTake2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    Combo(EditorIndex).Take_ReqItem2 = chkTake2.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "chkTake2_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbGSkill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    If Not frmEditor_Combos.Visible Then Exit Sub
    Combo(EditorIndex).GiveSkill = cmbGSkill.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbGSkill_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbItems1_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    If Not frmEditor_Combos.Visible Then Exit Sub
    Combo(EditorIndex).ReqItem1 = cmbItems1.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbItems1_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbItems2_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    If Not frmEditor_Combos.Visible Then Exit Sub
    Combo(EditorIndex).ReqItem2 = cmbItems2.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbItems2_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbSkill_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    If Not frmEditor_Combos.Visible Then Exit Sub
    Combo(EditorIndex).Skill = cmbSkill.ListIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmbSkill_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ComboEditorCancel
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdCancel_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    
    ClearCombo EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ":", EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ComboEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdDelete_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ComboEditorOk
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSave_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call ComboEditorOk(False)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "cmdSSave_Click", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "C1")
Me.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L1")
Frame3.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L2")
Frame1.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L3")
lblFirst.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L4")
chkItem1.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L5")
lblSecond.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L6")
chkItem2.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L7")
Frame4.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L8")
lblIndex.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L9")
lblNum.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L10")
lblValue.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L11")
lblUseless(0).Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L12")
lblSkillExp.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L13")
Frame2.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L14")
lblUseless(1).Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L15")
lblSLevel.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L16")
lblLevel.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L17")
lblUseless(2).Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L18")
lblItemVal1.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L19")
chkTake1.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L20")
lblUseless(3).Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L21")
lblItemVal2.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "L22")
chkTake2.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "B1")
cmdSSave.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "B2")
cmdDelete.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "B3")
cmdCancel.Caption = trad
trad = GetVar(App.Path & Lang, "CombEditor", "B4")
cmdSave.Caption = trad

Dim I As Long
    scrlItem1.max = MAX_ITEMS
    scrlItem2.max = MAX_ITEMS
    If cmbSkill.ListIndex > -1 Then
        scrlSkillLevel.max = Skill(cmbSkill.ListIndex + 1).MaxLvl
    Else
        scrlSkillLevel.max = Skill(1).MaxLvl
    End If
    scrlIndex.max = MAX_COMBO_GIVEN
    scrlLevel.max = MAX_LEVELS
    scrlGive.max = MAX_ITEMS
    scrlGiveVal.max = MAX_INTEGER
    scrlSkillExp.max = MAX_INTEGER
    scrlItemVal1.max = MAX_ITEMS
    scrlItemVal2.max = MAX_ITEMS
    COMBO_EDITOR_ITEM_INDEX = 1
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ComboEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Combo", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGive_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L9")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblNum.Caption = trad & " " & Trim$(Item(scrlGive.Value).name)
    Combo(EditorIndex).Item_Given(COMBO_EDITOR_ITEM_INDEX) = scrlGive.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlGive_Change", "frmEditor_Combo", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGiveVal_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L10")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblValue.Caption = trad & " " & scrlGiveVal.Value
    Combo(EditorIndex).Item_Given_Val(COMBO_EDITOR_ITEM_INDEX) = scrlGiveVal.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlGiveVal_Change", "frmEditor_Combo", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlIndex_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L8")
        ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblIndex.Caption = trad & " " & scrlIndex.Value
    COMBO_EDITOR_ITEM_INDEX = scrlIndex.Value
    ComboEditorInit
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlIndex_Change", "frmEditor_Combo", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem1_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L3")
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblFirst.Caption = trad & " " & Trim$(Item(scrlItem1.Value).name)
    Combo(EditorIndex).Item_1 = scrlItem1.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItem1_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem1_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlItem1.Value = 1 And scrlItem2.Value = 1 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    If Combo(EditorIndex).Item_2 > 0 Then
        lstIndex.AddItem EditorIndex & ": " & Trim$(Item(Combo(EditorIndex).Item_1).name) & " + " & Trim$(Item(Combo(EditorIndex).Item_2).name), EditorIndex - 1
    Else
        lstIndex.AddItem EditorIndex & ": " & Trim$(Item(Combo(EditorIndex).Item_1).name), EditorIndex - 1
    End If
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItem1_Validate", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem2_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L5")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblSecond.Caption = trad & " " & Trim$(Item(scrlItem2.Value).name)
    Combo(EditorIndex).Item_2 = scrlItem2.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItem2_Change", "frmEditor_Combo", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItem2_Validate(Cancel As Boolean)
Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlItem1.Value = 1 And scrlItem2.Value = 1 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Item(Combo(EditorIndex).Item_1).name) & " + " & Trim$(Item(Combo(EditorIndex).Item_2).name), EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItem2_Validate", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItemVal1_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L18")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblItemVal1.Caption = trad & " " & scrlItemVal1.Value
    Combo(EditorIndex).ReqItemVal1 = scrlItemVal1.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItemVal1_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlItemVal2_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L18")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblItemVal2.Caption = trad & " " & scrlItemVal2.Value
    Combo(EditorIndex).ReqItemVal2 = scrlItemVal2.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlItemVal2_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L16")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblLevel.Caption = trad & " " & scrlLevel.Value
    Combo(EditorIndex).Level = scrlLevel.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSkillExp_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L12")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblSkillExp.Caption = trad & " " & scrlSkillExp.Value
    Combo(EditorIndex).GiveSkill_Exp = scrlSkillExp.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlSkillExp_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSkillLevel_Change()
Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
Lang = "\data files\lang\" & Lang & ".ini"
Dim trad As String

trad = GetVar(App.Path & Lang, "CombEditor", "L15")
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_COMBO Then Exit Sub
    lblSLevel.Caption = trad & " " & scrlSkillLevel.Value
    Combo(EditorIndex).SkillLevel = scrlSkillLevel.Value
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "scrlLevel_Change", "frmEditor_Combos", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
