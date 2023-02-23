VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolbarButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Toolbar As Toolbar
Private m_Key     As String
Private WithEvents mImage As MSForms.Image
Attribute mImage.VB_VarHelpID = -1
Private Const MouseOverBorderColor As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private Const MouseOverBackColor   As Long = &HF3D7B3 ' = RGB(179, 215, 243)
Private Const MouseDownBorderColor As Long = &HD77800 ' = RGB(  0, 120, 215)
Private Const MouseDownBackColor   As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private m_IsMouseDown As Boolean
Private m_MouseDownX As Single
Private m_MouseDownY As Single

Friend Sub New_(aImage As MSForms.Image, aKey As String)
    Set mImage = aImage: m_Key = aKey
End Sub

Friend Property Set Owner(ByVal Value As Toolbar)
    Set m_Toolbar = Value
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Friend Sub Deselect()
    DrawButton_Clear
End Sub

Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = True
        m_MouseDownX = X
        m_MouseDownY = Y
        DrawButton_MouseDown
    End If
End Sub

Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_Toolbar.SelectButton Me
    If m_IsMouseDown And (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
        DrawButton_MouseDown
        Exit Sub
    End If
    DrawButton_MouseOver
End Sub

Private Sub mImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = False
        Deselect
        If (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
            m_Toolbar.OnClick Me
        End If
    End If
End Sub

Private Sub DrawButton_Clear()
    mImage.BorderStyle = fmBorderStyleNone
    mImage.BackStyle = fmBackStyleTransparent
End Sub

Private Sub DrawButton_MouseOver()
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseOverBorderColor ' &HEBBC80 'RGB(128, 188, 235)
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseOverBackColor     ' &HF3D7B3 'RGB(179, 215, 243)
End Sub

Private Sub DrawButton_MouseDown()
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BorderColor = MouseDownBorderColor ' &HD77800 'RGB(0, 120, 215)
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseDownBackColor     ' &HF3D7B3 'RGB(179, 215, 243)
End Sub


