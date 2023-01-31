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
Private Const MouseOverBackColor As Long = &HF3D7B3 ' = RGB(179, 215, 243)
Private Const MouseDownBackColor As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private Const ImageBorderColor   As Long = &HD77800 ' = RGB(0, 120, 215)
Private m_IsMouseDown As Boolean
Private m_IsMouseOver As Boolean
Private m_MouseDownX As Single
Private m_MouseDownY As Single

Friend Sub New_(aImage As MSForms.Image, aKey As String)
    Set mImage = aImage: m_Key = aKey
    mImage.BorderColor = ImageBorderColor ' &HD77800 'RGB(0, 120, 215)
End Sub

Friend Property Set Owner(ByVal Value As Toolbar)
    Set m_Toolbar = Value
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Sub Deselect()
    'm_IsMouseDown = False
    'm_IsMouseOver = False
    mImage.BorderStyle = fmBorderStyleNone
    mImage.BackStyle = fmBackStyleTransparent
    'UpdateView
End Sub

Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = True
        'UpdateView
        m_MouseDownX = X
        m_MouseDownY = Y
        mImage.BackStyle = fmBackStyleOpaque
        mImage.BackColor = MouseDownBackColor '&HEBBC80 'RGB(128, 188, 235)
    End If
End Sub

Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    m_Toolbar.SelectButton Me
    'm_IsMouseOver = True
    'UpdateView
    If m_IsMouseDown Then
        If (Abs(m_MouseDownX - X) > 25) Or (Abs(m_MouseDownY - Y) > 25) Then
            'm_IsMouseDown = False
        Else
            Exit Sub
        End If
    End If
    mImage.BorderStyle = fmBorderStyleSingle
    mImage.BackStyle = fmBackStyleOpaque
    mImage.BackColor = MouseOverBackColor '&HF3D7B3 'RGB(179, 215, 243)
End Sub

Private Sub mImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = XlMouseButton.xlPrimaryButton Then
        m_IsMouseDown = False
        Deselect
        If (Abs(m_MouseDownX - X) > 25) Or (Abs(m_MouseDownY - Y) > 25) Then Exit Sub
        m_Toolbar.OnClick Me
    End If
End Sub
'
'Private Sub UpdateView()
'    If m_IsMouseOver Then
'        mImage.BorderStyle = fmBorderStyleSingle
'        mImage.BackStyle = fmBackStyleOpaque
'        mImage.BackColor = MouseOverBackColor '&HF3D7B3 'RGB(179, 215, 243)
'    End If
'    If m_IsMouseDown Then
'        mImage.BackStyle = fmBackStyleOpaque
'        mImage.BackColor = MouseDownBackColor '&HEBBC80 'RGB(128, 188, 235)
'    ElseIf Not m_IsMouseOver Then
'        mImage.BorderStyle = fmBorderStyleNone
'        mImage.BackStyle = fmBackStyleTransparent
'    End If
'    'mImage.BorderColor = &H80000003  '&H80000002 'Inaktive Titelleiste
'    'mImage.BorderColor
'    'mImage.BorderStyle
'    'mImage.BackColor
'End Sub
