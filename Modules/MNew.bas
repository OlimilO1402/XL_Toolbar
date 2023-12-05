Attribute VB_Name = "MNew"
Option Explicit

#If VBA Then
Public Function Toolbar(Owner As UserForm, aFrame As MSForms.Frame) As Toolbar
#Else
Public Function Toolbar(Owner As Form, aFrame As Frame) As Toolbar
#End If
    Set Toolbar = New Toolbar: Toolbar.New_ Owner, aFrame
End Function

#If VBA Then
Public Function ToolbarButton(aImg As MSForms.Image, aKey As String) As ToolbarButton
#Else
Public Function ToolbarButton(aImg As PictureBox, aBorder As Shape, aKey As String) As ToolbarButton
#End If
    Set ToolbarButton = New ToolbarButton: ToolbarButton.New_ aImg, aBorder, aKey
End Function

#If VBA Then
Public Sub ShowUserForm()
#Else
Sub Main()
#End If
    UserForm1.Show
End Sub

