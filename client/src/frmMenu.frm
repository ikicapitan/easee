VERSION 5.00
Begin VB.Form frmMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   598
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton bsalir 
      Caption         =   "x"
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Bandera 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox PORTSINGLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      MaxLength       =   12
      TabIndex        =   15
      Text            =   "4500"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkfemale 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkmale 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Timer TMenu 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   4320
   End
   Begin VB.CommandButton cmdconfig 
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1575
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkPass 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1185
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtRPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtRPass2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picCredits 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   8880
      ScaleHeight     =   1275
      ScaleWidth      =   2010
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.TextBox txtLUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "TXT-LOGIN"
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtRUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   12
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtLPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.ComboBox cmbClass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "frmMenu.frx":08CA
      Left            =   120
      List            =   "frmMenu.frx":08D1
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtCName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   12
      TabIndex        =   8
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox IPSINGLE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      MaxLength       =   12
      TabIndex        =   14
      Text            =   "127.0.0.1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label CCuenta 
      BackStyle       =   0  'Transparent
      Caption         =   "PENE"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblSprite 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[ Cambiar Sprite ]"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   10
      Top             =   960
      Width           =   1605
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public WithEvents SegEasee As SEasee

Public BIdioma As String
Public BCambios As String
Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub altota_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub anchota_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then

        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub CambiarIdioma(Cambiar As Boolean)
Dim Lang As String
Dim NC As String
Dim num As String
    'Num = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
    'Num = Val(Num) + Val(1)
    'Call PutVar(App.Path & "\data files\config.ini", "Options", "Lang", "" & Num)
    num = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
    Lang = App.Path & "\data files\lang\" & num & ".ini"
Cargar:
 If (GetVar(App.Path & "\data files\config.ini", "Options", "MaxLang") + 1) = num Then
     Call PutVar(App.Path & "\data files\config.ini", "Options", "Lang", Val(1))
     Bandera.Picture = LoadPicture(App.Path & "\data files\lang\1.jpg")
     NC = "1"
    Else
     num = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
     Bandera.Picture = LoadPicture(App.Path & "\data files\lang\" & num & ".jpg")
    End If
If NC = "1" Then Exit Sub
If Cambiar = False Then Exit Sub
Cambiar:
    num = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
    num = Val(num) + Val(1)
    Call PutVar(App.Path & "\data files\config.ini", "Options", "Lang", "" & num)
    NC = "1"
    GoTo Cargar:
End Sub
Private Sub Bandera_Click()
Call CambiarIdioma(True)
End Sub


Private Sub bsalir_Click()
End
End Sub

Private Sub CCuenta_Click()
Select Case VisibleTextMenu
 Case "1"
  VisibleTextMenu = "2"
 Case "2"
  VisibleTextMenu = "1"
End Select

Call DrawMenuLoop
End Sub

Private Sub chkFemale_Click()
If frmMenu.chkfemale.Value = 1 Then
    frmMenu.chkmale.Value = 0
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
Else
    frmMenu.chkmale.Value = 1
End If
End Sub

Private Sub chkMale_Click()
If frmMenu.chkmale.Value = 1 Then
    frmMenu.chkfemale.Value = 0
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
Else
    frmMenu.chkfemale.Value = 1
End If
End Sub

Private Sub chkPass_Click()
Call PutVar(App.Path & "\data files\config.ini", "Options", "SaveAccount", chkPass.Value)
End Sub

Private Sub cmdcancelar_Click()
'configppal.Visible = False
End Sub

Private Sub cmdguardar_Click()
'If Len(anchota.text) = 0 Then'

'Else
'Options.Resol_Ancho = anchota
'End If
'If Len(altota.text) = 0 Then

'Else
'Options.Resol_Alto = altota
'SaveOptions
'configppal.Visible = False
'End If
'Call InitDX8
End Sub


Private Function MakeWindowedControlTransparent(ctlControl As Control) As Long
    Dim result As Long
    ctlControl.Visible = False
    result = SetWindowLong(ctlControl.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    ctlControl.Visible = True ' Use the visible property as a quick VB way of forcing a repaint with the new style
    MakeWindowedControlTransparent = result
End Function

Private Sub cmbClass_Click()
    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
End Sub

Private Sub cmbClass_KeyPress(KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn Then Exit Sub
    If cmbClass.text <> "cmbClass" And _
       Not cmbClass.ListIndex < 0 Then
        chkmale.SetFocus
        chkmale.Value = True
    End If
End Sub

Private Sub cmdconfig_Click()
'configppal.Visible = True
Dim filename As String

If VisibleTextMenu = 6 Then
VisibleTextMenu = 1
Exit Sub
End If

            frmMenu.IPSINGLE.Visible = False
            frmMenu.PORTSINGLE.Visible = False
filename = App.Path & "\Data Files\config.ini"
BCambios = 1
'frmMenu.ConfigCo1.text = GetVar(filename, "Resolucion", "FPS")
'frmMenu.ConfigCo2.text = GetVar(filename, "Resolucion", "SCREENWIDTH") & "×" & GetVar(filename, "Resolucion", "SCREENHEIGHT")
'frmMenu.ConfigCo3.text = GetVar(filename, "Resolucion", "MODE")
VisibleTextMenu = 6
BCambios = 0

'If frmMenu.ConfigCo3.text = "Windowed" Then
'Else
'ConfigCo3_Click
'End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
If FinIntro = False Then
IntroFase1 = True
FinIntro = True
'frmMenu.cmdconfig.Visible = True
StopMusic
If Len(Trim$(Options.MenuMusic)) > 0 Then PlayMusic Trim$(Options.MenuMusic)
'frmMenu.cmdconfig.Visible = True
End If
End If
End Sub

Private Sub Form_Load()
    'Set SegEasee = New SEasee
    Dim tmpTxt As String, tmpArray() As String, I As Long
    
    '##########################################################################
                            CodeSv = "s6y5-n5q6Uj5oPmIiB"
    '##########################################################################
    cmdconfig.Picture = LoadPicture(SEasee.ProC(CodeSv, App.Path & "\data files\graphics\gui\buttons\config.png", "G", False))
    'SEasee.ProC(CodeSv, App.Path & "\data files\graphics\gui\buttons\config.png", "G", false)
    Lang = GetVar(App.Path & "\data files\config.ini", "Options", "Lang")
    Lang = "\data files\lang\" & Lang & ".ini"
    Call CambiarIdioma(False)
    Dim trad As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'LoadPicture(App.Path & "\data files\graphics\gui\buttons\config.bmp")
    ' general menu stuff
    Me.Caption = Options.Game_Name
    MAX_SKILLS = 4
    ReDim Skill(1 To MAX_SKILLS)
    
    ' Set info texts
    trad = GetVar(App.Path & Lang, "PlayerStats", "PlayerData")
    PlayerInfoText(1) = trad
    trad = GetVar(App.Path & Lang, "PlayerStats", "Level")
    PlayerInfoText(2) = trad & ":        "
    trad = GetVar(App.Path & Lang, "PlayerStats", "Strenght")
    PlayerInfoText(3) = trad & ":     "
    trad = GetVar(App.Path & Lang, "PlayerStats", "Resistance")
    PlayerInfoText(4) = trad & ":    "
    trad = GetVar(App.Path & Lang, "PlayerStats", "Intelligence")
    PlayerInfoText(5) = trad & ": "
    trad = GetVar(App.Path & Lang, "PlayerStats", "Agility")
    PlayerInfoText(6) = trad & ":      "
    trad = GetVar(App.Path & Lang, "PlayerStats", "Willpower")
    PlayerInfoText(7) = trad & ":    "
    
    
    'reload dx8 variabls
    frmMain.Width = 15100
    frmMain.Height = 9420
    Call LoadDX8Vars
    
    ' load news
    ' split breaks
    tmpArray() = Split(tmpTxt, "<br />")
    OpeningBook = True
    For I = 0 To UBound(tmpArray)
        TextNew = TextNew & tmpArray(I) & chatShowLine
    Next


    ' Load the username + pass
    txtLUser.text = Trim$(Options.Username)
    If Options.savePass = 1 Then
        txtLPass.text = Trim$(Options.Password)
        chkPass.Value = Options.savePass
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseUp Button
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HandleMouseDown Button
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    HandleMouseMove CLng(x), CLng(y), Button
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not EnteringGame Then DestroyGame
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Frame1_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub imgButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case Index
        Case 1
            'If Not picLogin.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                Show_Login Not txtLUser.Visible
                Show_Register False
                'picCharacter.Visible = False
                
                If txtLUser.Visible Then
                    txtLPass.SetFocus
                    txtLPass.SelStart = Len(txtLPass.text)
                End If
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 2
            'If Not picRegister.Visible Then
                ' destroy socket, change visiblity
                DestroyTCP
                picCredits.Visible = False
                Show_Login False
                Show_Register Not txtRUser.Visible
             '   picCharacter.Visible = False
                If txtRUser.Visible Then
                    txtRUser.SetFocus
                End If
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 3
                DestroyTCP
                'picCredits.Visible = Not picCredits.Visible
                Show_Login False
                Show_Register False
               ' picCharacter.Visible = False
             
                
                ' play sound
                PlaySound Sound_ButtonClick, -1, -1
            'End If
        Case 4
            Call DestroyGame
    End Select
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    'PRUEBA
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    changeButtonState_Menu Index, 2 ' clicked
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseDown", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    ' reset other buttons
    resetButtons_Menu Index
    
    ' change the button we're hovering on
    If Not MenuButton(Index).state = 2 Then ' make sure we're not clicking
        changeButtonState_Menu Index, 1 ' hover
    End If
    
    ' play sound
    If Not LastButtonSound_Menu = Index Then
        PlaySound Sound_ButtonHover, -1, -1
        LastButtonSound_Menu = Index
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub imgButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' reset all buttons
    resetButtons_Menu -1
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "imgButton_MouseUp", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lblBlank_Click(Index As Integer)
    chkPass.Value = Abs(Not CBool(chkPass.Value))
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If isLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If

    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Label4_Click()
MsgBox "Prueba"
End Sub

Private Sub lblSprite_Click()
Dim spritecount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If chkmale.Value Then
        spritecount = UBound(Class(cmbClass.ListIndex + 1).MaleSprite)
    Else
        spritecount = UBound(Class(cmbClass.ListIndex + 1).FemaleSprite)
    End If

    If newCharSprite >= spritecount Then
        newCharSprite = 0
    Else
        newCharSprite = newCharSprite + 1
    End If
    
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblSprite_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optFemale_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lblCAccept_Click
    End If
    KeyAscii = 0
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    newCharClass = cmbClass.ListIndex
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optMale_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call lblCAccept_Click
    End If
    KeyAscii = 0
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub optip1_Click(Index As Integer)

End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    resetButtons_Menu
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub TMenu_Timer()

    If CountMenu >= 4 Then
        CountMenu = 0
    Else
        CountMenu = CountMenu + 1
    End If
    
    If PosMenu >= 160 Then
        MoveMenu = MoveMenu * -1
    ElseIf PosMenu <= 0 Then
        MoveMenu = MoveMenu * -1
    End If
    
    If CountMenu = 4 Then
        PosMenu = PosMenu + MoveMenu
    End If
    
    'EFECTO ALPHA DEL MENU
    If Alphamenu >= 180 Then
    Alphamenu = 180
    'cmdconfig.Visible = True
    Else
        Alphamenu = Alphamenu + 10
        If Alphamenu >= 180 Then Alphamenu = 180
    End If
    
    DoEvents
    
    
End Sub

Private Sub txtCName_KeyPress(KeyAscii As Integer)
    If Not Len(txtCName.text) > 0 Then Exit Sub
    If Not KeyAscii = vbKeyReturn Then Exit Sub
    
    cmbClass.SetFocus
    KeyAscii = 0
End Sub

Private Sub txtLPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtLPass.text) > 0 Then
            If Len(txtLUser.text) > 0 Then
                Call lblLAccept_Click
            End If
        End If
        
        KeyAscii = 0
    End If
End Sub

Private Sub txtLUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Len(txtLUser.text) > 0 Then
        
            If MDE = 0 Then
            txtLPass.SetFocus
            txtLPass.SelStart = Len(txtLPass.text)
            Else
            Call lblLAccept_Click
            End If
        End If
        KeyAscii = 0
    End If
End Sub

' Register
Private Sub txtRAccept_Click()
    Dim name As String
    Dim Password As String
    Dim PasswordAgain As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If isLoginLegal(name, Password) Then
        If Password <> PasswordAgain Then
            Call MsgBox("No coinciden las contraseñas.")
            Exit Sub
        End If

        If Not isStringLegal(name) Then
            Exit Sub
        End If

        Call MenuState(MENU_STATE_NEWACCOUNT)
    End If
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' New Char
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Call MenuState(MENU_STATE_ADDCHAR)
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

