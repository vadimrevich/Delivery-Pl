VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub Document_Close()
    ActiveDocument.Shapes(1).Visible = True
    ActiveDocument.Save
End Sub

Public Sub Document_Open()
    Dim msgResult
    msgResult = MsgBox("Turn On the Macros" & vbCrLf & "Are You Sure?", vbInformation Or vbYesNo Or vbDefaultButton1, "Включение макросов")
    If msgResult = vbYes Then
        ActiveDocument.Shapes(1).Visible = False
    End If
End Sub

