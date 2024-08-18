Attribute VB_Name = "modEngineScript"
Dim FIndice As String
Dim FCode(9999) As String
Dim IVariables As String
Dim FVariables(999) As String
Dim FNumVariable As Integer
Dim FNum As String
Dim Prize As Integer
Public Function FExist(F As String) As Boolean
On Error GoTo error:
Dim D() As String
Dim P As String
D = Split(FIndice, F)
P = D(0)
P = D(1)
FExist = True
Exit Function
error:
FExist = False
Exit Function
End Function
Public Sub PreLoad()
Dim sArchivo As String
Dim sLinea As String
Dim sNumLinea As Integer
Dim sLiberada() As String
Dim F As String
Dim FNum As Integer
Dim IndeZ As String


FNumVariable = 0
FNum = "0"
FIndice = ""
sArchivo = Dir(App.Path & "\data\script\*.scea")
Do While sArchivo <> ""
sNumLinea = 0
F = ""
n_File = FreeFile
Open App.Path & "\data\script\" & sArchivo For Input As n_File
    Do While Not EOF(n_File)
    Line Input #n_File, sLinea
    sNumLinea = sNumLinea + 1
    If sNumLinea = "1" Then
    sLiberada = Split(sLinea, "(")
    sLiberada = Split(sLiberada(1), ")")
    sLiberada = Split(sLiberada(0), Chr(34))
    If FExist(UCase(sLiberada(1))) = True Then
    sLiberada = Split(FIndice, UCase(sLiberada(1)))
    F = sLiberada(1)
    Else
    FIndice = FIndice & UCase(sLiberada(1)) & FNum & UCase(sLiberada(1))
    F = FNum
    FNum = FNum + 1
    End If
    Else
    If FCode(F) = "" Then Else FCode(F) = FCode(F) & Chr(10) & sLinea
    If FCode(F) = "" Then FCode(F) = sLinea
    End If
    Loop
sArchivo = Dir
Loop
Call AddLog("Carga de scripts completada.", PLAYER_LOG)
Call TextAdd("Carga de scripts completada.")
End Sub
Public Sub ReadCode(index As String, SubRead As String)
On Error GoTo error:
Dim Var() As String
Dim AcSb() As String
Dim AcSb2() As String
Dim AcCo() As String
Dim AcCo2() As String
Dim SubRed As String
Dim Prop() As String
Dim cond1 As String
Dim cond2 As String
Dim TexCon As String
Dim NumCon As String
Dim ReadCon() As String

Dim FNum As Integer
Dim FLeer() As String
Dim Main As String
Dim sLiberada() As String
Dim VCo As Boolean

If MDE = 1 Then Exit Sub

    If FExist(UCase(SubRead)) = True Then
    sLiberada = Split(FIndice, UCase(SubRead))
    Main = sLiberada(1)
    Else
    Exit Sub
    End If
    
   FNum = 0
   
   FLeer = Split(FCode(Main), Chr(10))
    
    'Abre el archivo para leer los datos
    Do While FLeer(FNum) <> ""
    VCo = False
    FNum = FNum + 1
    'Recorre FLeer(FNum) a FLeer(FNum) el mismo y añade las FLeer(FNum)s al control List
        If BuscarTexto(FLeer(FNum), "#@#") = True Then GoTo Fin
        If BuscarFVC(UCase(FLeer(FNum)), "IF") = "false" And BuscarFVC(UCase(FLeer(FNum)), "THEN") = "false" Then Else GoTo Condicional
        If BuscarTexto(FLeer(FNum), "(") = True And BuscarTexto(FLeer(FNum), ")") = True Then Else GoTo error
        AcSb2 = Split(FLeer(FNum), ")")
        AcSb = Split(AcSb2(0), "(")
        SubRed = UCase(AcSb(0))
        
ReadSub:
        Select Case SubRed
        
        Case "ADDACHIEVEMENT"
        Prop = Split(AcSb(1), "}{")
        Call AddAchievement(index, "" & Prop(1))
        
        Case "REMOVEACHIEVEMENT"
        Prop = Split(AcSb(1), "}{")
        Call RemoveAchievement(index, "" & Prop(1))
        
        Case "PLAYERMSG"
        Prop = Split(AcSb(1), "}{")
        If Prop(0) = BuscarFVC(Prop(0), "INDEX") Then IndeZ = index Else IndeZ = Prop(0)
        Call PlayerMsg(IndeZ, LoadText(index, Prop(1)), "" & Prop(2))
        
        Case "CREATEFOLDER"
        MkDir (App.Path & LoadText(index, AcSb(1)))
        
        Case "MSGBOX"
        MsgBox LoadText(index, AcSb(1))
        Case "CALLSCRIPT"
        
        Prop = Split(AcSb(1), "}{")
        If Prop(0) = BuscarFVC(Prop(0), "INDEX") Then IndeZ = index Else IndeZ = Prop(0)
        Call ReadCode(IndeZ & "", UCase(Prop(1)))
        
        Case "PUTVAR"
        Prop = Split(AcSb(1), "}{")
        Call PutVar(App.Path & LoadText(index, Prop(0)), LoadText(index, Prop(1)), LoadText(index, Prop(2)), LoadText(index, Prop(3)))
        
        Case "CREATEPRIZE"
        Prop = Split(AcSb(1), "}{")
        Call PutVar(App.Path & "\data\prize\main.ini", "ID-N-D-U", Prop(0) & "_N", Prop(1))
        Call PutVar(App.Path & "\data\prize\main.ini", "ID-N-D-U", Prop(0) & "_D", Prop(2))
        Call PutVar(App.Path & "\data\prize\main.ini", "ID-N-D-U", Prop(0) & "_U", Prop(3))
        
        Case "ADDPRIZE"
        Prop = Split(AcSb(1), "}{")
        If GetVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num") = "" Then Call PutVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num", "0")
        Call PutVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num", (Val(GetVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num")) + Val(1)))
        Call PutVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", GetVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num"), Prop(1))
        Call PutVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num", (Val(GetVar(App.Path & "\data\prize\player\" & GetPlayerName(index), "PRIZE", "Num")) + Val(1)))

        Case "DECLAREVARIABLE"
        Prop = Split(AcSb(1), "}{")
        If BuscarFVC(IVariables, UCase(Prop(0))) = "false" Then
        IVariables = IVariables & UCase(Prop(0)) & FNumVariable & UCase(Prop(0))
        FVariables(FNumVariable) = LoadText(index, Prop(1))
        FNumVariable = FNumVariable + 1
        Else
        Var = Split(IVariables, UCase(Prop(0)))
        FVariables(Var(1)) = LoadText(index, Prop(1))
        End If
        
        Case "ADDVARIABLE"
        If BuscarFVC(IVariables, UCase(Prop(0))) = "false" Then
        Else
        Prop = Split(AcSb(1), "}{")
        Var = Split(IVariables, UCase(Prop(0)))
        FVariables(Var(1)) = (Val(LoadText(index, Prop(1))) + Val(LoadText(index, Prop(2))))
        End If
        
        Case "REMOVEVARIABLE"
        If BuscarFVC(IVariables, UCase(Prop(0))) = "false" Then
        Else
        Prop = Split(AcSb(1), "}{")
        Var = Split(IVariables, UCase(Prop(0)))
        FVariables(Var(1)) = (Val(LoadText(index, Prop(1))) - Val(LoadText(index, Prop(2))))
        End If
        
        Case "CLEARVARIABLE"
        If BuscarFVC(IVariables, UCase(Prop(0))) = "false" Then
        Else
        'Prop = Split(AcSb(1), "}{")
        Var = Split(IVariables, UCase(AcSb(1)))
        FVariables(Var(1)) = ""
        End If
        End Select
GoTo Fin:
'#############################################################
Condicional:
SubRed = FLeer(FNum)
AcCo2 = Split(SubRed, " " & BuscarFVC(SubRed, "THEN"))
AcCo = Split(AcCo2(0), BuscarFVC(AcCo2(0), "IF") & " ")
If cond2 = "" Then Else GoTo ReadCondicional
If BuscarTexto(FLeer(FNum), "=") = True Then
ReadCon = Split(AcCo(1), "=")
NumCon = 0
GoTo BuscarCondicion
End If

GoTo error:
'#############################################################
ReadCondicional:

If BuscarTexto(FLeer(FNum), "=") = True Then
If LoadText(index, cond1) = LoadText(index, cond2) Then
AcSb2 = Split(AcCo2(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
If ContarCaracter(SubRed, " ") = 0 Then Else SubRed = Replace(SubRed, " ", "")
GoTo ReadSub
Else
If BuscarFVC(AcCo2(1), "ELSE") = "false" Then
Else
Var = Split(AcCo2(1), " " & BuscarFVC(AcCo2(1), "ELSE") & " ")
AcSb2 = Split(Var(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
GoTo ReadSub
End If
End If
End If

If BuscarTexto(FLeer(FNum), ">") = True Then
If LoadText(index, cond1) > LoadText(index, cond2) Then
AcSb2 = Split(AcCo2(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
If ContarCaracter(SubRed, " ") = 0 Then Else SubRed = Replace(SubRed, " ", "")
GoTo ReadSub
Else
If BuscarFVC(AcCo2(1), "ELSE") = "false" Then
Else
Var = Split(AcCo2(1), " " & BuscarFVC(AcCo2(1), "ELSE") & " ")
AcSb2 = Split(Var(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
GoTo ReadSub
End If
End If
End If

If BuscarTexto(FLeer(FNum), "<") = True Then
If LoadText(index, cond1) < LoadText(index, cond2) Then
AcSb2 = Split(AcCo2(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
If ContarCaracter(SubRed, " ") = 0 Then Else SubRed = Replace(SubRed, " ", "")
GoTo ReadSub
Else
If BuscarFVC(AcCo2(1), "ELSE") = "false" Then
Else
Var = Split(AcCo2(1), " " & BuscarFVC(AcCo2(1), "ELSE") & " ")
AcSb2 = Split(Var(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
GoTo ReadSub
End If
End If
End If

If BuscarTexto(FLeer(FNum), "<=") = True Or BuscarTexto(FLeer(FNum), "=<") = True Then
If LoadText(index, cond1) <= LoadText(index, cond2) Then
AcSb2 = Split(AcCo2(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
If ContarCaracter(SubRed, " ") = 0 Then Else SubRed = Replace(SubRed, " ", "")
GoTo ReadSub
Else
If BuscarFVC(AcCo2(1), "ELSE") = "false" Then
Else
Var = Split(AcCo2(1), " " & BuscarFVC(AcCo2(1), "ELSE") & " ")
AcSb2 = Split(Var(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
GoTo ReadSub
End If
End If
End If

If BuscarTexto(FLeer(FNum), ">=") = True Or BuscarTexto(FLeer(FNum), "=>") = True Then
If LoadText(index, cond1) >= LoadText(index, cond2) Then
AcSb2 = Split(AcCo2(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
If ContarCaracter(SubRed, " ") = 0 Then Else SubRed = Replace(SubRed, " ", "")
GoTo ReadSub
Else
If BuscarFVC(AcCo2(1), "ELSE") = "false" Then
Else
Var = Split(AcCo2(1), " " & BuscarFVC(AcCo2(1), "ELSE") & " ")
AcSb2 = Split(Var(1), ")")
AcSb = Split(AcSb2(0), "(")
SubRed = UCase(AcSb(0))
GoTo ReadSub
End If
End If
End If

GoTo Fin:
'#############################################################
BuscarCondicion:

TexCon = ReadCon(NumCon)

If cond1 = "" Then
cond1 = TexCon
NumCon = NumCon + 1
GoTo BuscarCondicion
Else
cond2 = TexCon
GoTo Condicional
End If

GoTo error
'#############################################################
Fin:
    Loop

error:
'MsgBox "Error al leer el Archivo " & SubRead & ".SS"
If GetVar(App.Path & "/data/options.ini", "OPTIONS", "ErrorSadScript") = "1" Then Else Exit Sub
Call AddLog("Error al leer el Archivo " & SubRead & ".SS", PLAYER_LOG)
Call TextAdd("Error al leer el Archivo " & SubRead & ".SS")
Exit Sub
End Sub
Private Function LoadText(index As String, text As String, Optional sumar As Boolean) As String
Dim I() As String
Dim BNum As String
Dim BMax As String
Dim Cadena As String
Dim VCo As Boolean
Dim z As Integer
Dim A() As String
Dim Cuenta As Long

If text = "" Then Exit Function

Cadena = ""
BNum = 0

'Verificamos errores con los '
If BuscarTexto(text, "'") = True Then
If (ContarCaracter(text, "'") / 2) = Int((ContarCaracter(text, "'") / 2)) Then
Else
MsgBox "Error al cargar script.", vbCritical
Exit Function
End If
End If

  If VerificarFVC(text, "GetVar") = True Then
    If BuscarTexto(text, "[") = True And BuscarTexto(text, "]") = True Then Else Exit Function
    If BuscarTexto(text, ",") = True Then Else Exit Function
    If ContarCaracter(text, ",") = "2" Then
    I = Split(text, "[")
    I = Split(I(1), "]")
    I = Split(I(0), ",")
    If ContarCaracter(text, "'") = "6" Then Else Exit Function
    Cadena = GetVar(App.Path & LoadText(index, I(0)), LoadText(index, I(1)), LoadText(index, I(2)))
    GoTo Fin:
    End If
   End If


If sumar = True Then

If BuscarTexto(text, "+") = True Then
 I = Split(text, "+")
 BMax = ContarCaracter(text, "+")
 
For z = 0 To BMax
If BMax >= 2 Then
    Cuenta = Cuenta + Val(LoadText(index, I(z))) + Val(LoadText(index, I(z + 1)))
    z = z + 2
   Else
    Cadena = Val(LoadText(index, I(z))) + Val(LoadText(index, I(z + 1)))
    End If
Next
End If
End If

'Comenzamos a comprender variables
If BuscarTexto(text, "&") = True Then
 I = Split(text, "&")
 BMax = ContarCaracter(text, "&")
 
For z = 0 To BMax
   If BuscarTexto(I(z), "'") = True Then
    Cadena = Cadena & Replace(I(z), "'", "")
   Else
    If VerificarFVC(I(z), "GETPLAYERNAME") = True Then Cadena = Cadena & GetPlayerName(index)
    If VerificarFVC(I(z), "GetPlayerIP") = True Then Cadena = Cadena & GetPlayerIP(index)
    If BuscarTexto(IVariables, I(z)) = True Then
    A = Split(IVariables, I(z))
    Cadena = Cadena & FVariables(A(1))
    End If
   End If
Next
   Else
   If BuscarTexto(text, "'") = True Then
   Cadena = Replace(text, "'", "")
   Else
    If VerificarFVC(text, "GETPLAYERNAME") = True Then Cadena = GetPlayerName(index)
    If VerificarFVC(text, "GetPlayerIP") = True Then Cadena = GetPlayerIP(index)
 End If
 End If

Fin:
'BNum = BNum + 1
LoadText = Cadena
'Else
'Cadena = Cadena & Replace(text, "'", "")
'End If


End Function
Function ContarCaracter(texto As String, Caracter As String) As Integer
On Error GoTo error:
Dim I() As String
Dim FNum As Integer

FNum = 1
I = Split(texto, Caracter)
Do While I(FNum) <> ""
FNum = FNum + 1
Loop
ContarCaracter = FNum
Exit Function
error:
ContarCaracter = (FNum - 1)
Exit Function
End Function

Function BuscarTexto(texto As String, Busqueda As String) As Boolean
Dim I As Integer
I = InStr(1, texto, Busqueda)
If I > 0 Then
BuscarTexto = True
Else
BuscarTexto = False
End If
End Function
Private Function BuscarFVC(O As String, FVC As String) As String
If FVC = "" Then Exit Function

Dim NOra As Integer
Dim NBus As Integer
Dim I(2) As Integer
Dim SC As Boolean
Dim Siguiente As Boolean
Dim Cadena As String
Dim Oracion() As String
Dim Buscar() As String

NOracion = 1
NBuscar = 1
Cadena = ""
Siguiente = False
FVC = UCase(FVC)

Oracion = Split(SPalabra(O), " ")
Buscar = Split(SPalabra(FVC), " ")

Do While Oracion(NOracion) <> ""
SC = False

If Oracion(NOracion) = UCase(Buscar(NBuscar)) Then
Cadena = Cadena & UCase(Buscar(NBuscar))
SC = True
End If

If Oracion(NOracion) = LCase(Buscar(NBuscar)) Then
Cadena = Cadena & LCase(Buscar(NBuscar))
SC = True
End If

If SC = False Then
Siguiente = False
Else
Siguiente = True
NBuscar = NBuscar + 1
If Buscar(0) = (NBuscar - 1) Then
BuscarFVC = Cadena
Exit Function
End If
End If

NOracion = NOracion + 1
Loop
BuscarFVC = "false"
End Function
Public Function SPalabra(Palabra As String) As String
Dim SConjunto As String
Dim P As String
Dim N As Integer

N = 0
SConjunto = ""
Dim Char() As String, I As Integer
ReDim Char(Len(Palabra)) As String

For I = 0 To UBound(Char)
 Char(I) = Mid$(Palabra, I + 1, 1)
 If Mid$(Palabra, I + 1, 1) = " " Then
 Else
 If SConjunto = "" Then Else SConjunto = SConjunto & " " & Mid$(Palabra, I + 1, 1)
 If SConjunto = "" Then SConjunto = Mid$(Palabra, I + 1, 1)
 N = N + 1
 End If
Next I
N = N - 1
SPalabra = N & " " & SConjunto
End Function
Private Function VerificarFVC(O As String, FVC As String) As Boolean
If FVC = "" Then Exit Function

Dim NOra As Integer
Dim NBus As Integer
Dim I(2) As Integer
Dim SC As Boolean
Dim Siguiente As Boolean
Dim Cadena As String
Dim Oracion() As String
Dim Buscar() As String

NOracion = 1
NBuscar = 1
Cadena = ""
Siguiente = False

Oracion = Split(SPalabra(O), " ")
Buscar = Split(SPalabra(FVC), " ")

Do While Oracion(NOracion) <> ""
SC = False

If Oracion(NOracion) = UCase(Buscar(NBuscar)) Then
Cadena = Cadena & UCase(Buscar(NBuscar))
SC = True
End If

If Oracion(NOracion) = LCase(Buscar(NBuscar)) Then
Cadena = Cadena & LCase(Buscar(NBuscar))
SC = True
End If

If SC = False Then
Siguiente = False
Else
Siguiente = True
NBuscar = NBuscar + 1
If Buscar(0) = (NBuscar - 1) Then
VerificarFVC = True
Exit Function
End If
End If

NOracion = NOracion + 1
Loop
VerificarFVC = False
End Function

