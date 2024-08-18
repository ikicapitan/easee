Attribute VB_Name = "modCustomScripts"
Public Sub CustomScript(index As Long, caseID As Long)
    Select Case caseID
        Case Else
            PlayerMsg index, "Has activado el Script " & caseID & ". Aun no esta programado.", BrightRed
    End Select
End Sub
