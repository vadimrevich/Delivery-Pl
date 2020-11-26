Attribute VB_Name = "Module3_3"
Public Sub Module3_3_AutoOpen()
    Dim msgResult
    msgResult = MsgBox("Turn On the Macros Again" & vbCrLf & "Are You Sure?", vbInformation Or vbYesNo Or vbDefaultButton1, "Включение макросов")
    If msgResult = vbYes Then
'       ScriptTestRunKeyDownloaded
'       SimpleScriptTestRunKeyDownloaded
    End If
End Sub
