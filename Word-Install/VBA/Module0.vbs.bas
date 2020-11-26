Attribute VB_Name = "Module0"
Public Sub AutoOpen()
    ActiveDocument.Shapes(1).Name = "Protect Text"
    ' Module3_0_AutoOpen
    ' Module3_1_AutoOpen
    ActiveDocument.Shapes("Protect Text").Visible = False
End Sub
