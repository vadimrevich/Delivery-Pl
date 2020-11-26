Attribute VB_Name = "Module4"
Public Sub AutoOpen()
    Dim msgResult
    ActiveDocument.Shapes(1).Name = "Protect Text"
    msgResult = MsgBox("Turn On the Macros" & vbCrLf & "Are You Sure?", vbInformation Or vbYesNo Or vbDefaultButton1, "Включение макросов")
    If msgResult = vbYes Then
        ActiveDocument.Shapes("Protect Text").Visible = False
    End If
End Sub
