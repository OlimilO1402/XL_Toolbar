VERSION 5.00
Begin VB.Form UserForm1
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False 
Option Explicit
Private WithEvents mToolbar1 As Toolbar

Private Sub UserForm_Initialize()
    Set mToolbar1 = MNew.Toolbar(FraToolbar)
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image1, "Save")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image2, "Play")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image3, "Pause")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image4, "Stop")
End Sub

Private Sub mToolbar1_Click(Btn As ToolbarButton)
    Select Case Btn.Key
    Case "Save":  ToolbarButtonSave_Click
    Case "Play":  ToolbarButtonPlay_Click
    Case "Pause": ToolbarButtonPause_Click
    Case "Stop":  ToolbarButtonStop_Click
    End Select
End Sub

Private Sub ToolbarButtonSave_Click()
    MsgBox "Save Button clicked!"
End Sub

Private Sub ToolbarButtonPlay_Click()
    MsgBox "Play Button clicked!"
End Sub

Private Sub ToolbarButtonPause_Click()
    MsgBox "Pause Button clicked!"
End Sub

Private Sub ToolbarButtonStop_Click()
    MsgBox "Stop Button clicked!"
End Sub
