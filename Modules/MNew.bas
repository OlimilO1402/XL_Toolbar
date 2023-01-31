Option Explicit

Public Function Toolbar(aFrame As MSForms.Frame) As Toolbar
    Set Toolbar = New Toolbar: Toolbar.New_ aFrame
End Function

Public Function ToolbarButton(aImg As MSForms.Image, aKey As String) As ToolbarButton
    Set ToolbarButton = New ToolbarButton: ToolbarButton.New_ aImg, aKey
End Function

Public Sub ShowUserForm()
    UserForm1.Show
End Sub
