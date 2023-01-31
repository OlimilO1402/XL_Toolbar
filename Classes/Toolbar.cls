VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Toolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Buttons As Collection
Private WithEvents mPanel   As MSForms.Frame
Public Event Click(Btn As ToolbarButton)

Private Sub Class_Initialize()
    Set m_Buttons = New Collection
End Sub

Friend Sub New_(Frame As MSForms.Frame)
    Set mPanel = Frame
End Sub

Public Sub AddButton(aBtn As ToolbarButton)
    m_Buttons.Add aBtn
    Set aBtn.Owner = Me
End Sub

Public Sub OnClick(Btn As ToolbarButton)
    RaiseEvent Click(Btn)
End Sub

Public Sub SelectButton(aBtn As ToolbarButton)
    Dim v, tbb As ToolbarButton
    For Each v In m_Buttons
        Set tbb = v
        If Not tbb Is aBtn Then
            tbb.Deselect
        End If
    Next
End Sub

Private Sub mPanel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SelectButton Nothing
End Sub
