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
#If VBA Then
    Private mOwner            As MSForms.UserForm
    Private WithEvents mPanel As MSForms.Frame
Attribute mPanel.VB_VarHelpID = -1
#Else
    Private mOwner              As Form
    Private WithEvents mPanel   As Frame
Attribute mPanel.VB_VarHelpID = -1
    Private WithEvents mBackImg As Image
Attribute mBackImg.VB_VarHelpID = -1
#End If
Private m_Buttons  As Collection
Public Event Click(Btn As ToolbarButton)

Private Sub Class_Initialize()
    Set m_Buttons = New Collection
End Sub

#If VBA Then
Friend Sub New_(Owner As UserForm, aFrame As MSForms.Frame)
#Else
Friend Sub New_(Owner As Form, aFrame As Frame)
#End If
    Set mOwner = Owner: Set mPanel = aFrame
End Sub

Public Sub AddButton(aBtn As ToolbarButton)
    m_Buttons.Add aBtn
    Set aBtn.Owner = Me
End Sub

Public Sub OnClick(Btn As ToolbarButton)
    RaiseEvent Click(Btn)
End Sub

Public Sub SelectButton(aBtn As ToolbarButton)
    Dim v, obj As ToolbarButton
    For Each v In m_Buttons
        Set obj = v
        If Not obj Is aBtn Then
            obj.Deselect
        End If
    Next
End Sub

Public Sub DeselectAll()
    SelectButton Nothing
End Sub

#If VBA Then
Private Sub mPanel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#Else
Private Sub mPanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
#End If
    SelectButton Nothing
End Sub

