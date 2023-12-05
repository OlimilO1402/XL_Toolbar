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
Private Enum EMouseButtons
    MButtonNone = 0    ' 0                                     'XlMouseButton.xlNoButton
    MButtonLeft = 1    ' 1 MouseButtonConstants.vbLeftButton   'XlMouseButton.xlPrimaryButton
    MButtonRight = 2   ' 2 MouseButtonConstants.vbRightButton  'XlMouseButton.xlSecondaryButton
    MButtonMiddle = 4  ' 4 MouseButtonConstants.vbMiddleButton
End Enum
Private Enum EBorderStyle
    BorderStyleNone = 0   ' 0 BorderStyleConstants 'fmBorderStyle.fmBorderStyleNone
    BorderStyleSingle = 1 ' 1                       'fmBorderStyle.fmBorderStyleSingle
    'BStyle =
End Enum
Private Enum EBackStyle
    BackStyleTransparent = 0 ' fmBackStyle.fmBackStyleTransparent
    BackStyleOpaque = 1      ' fmBackStyle.fmBackStyleOpaque
End Enum
#If VBA Then
    Private WithEvents mImage As MSForms.Image
Attribute mImage.VB_VarHelpID = -1
#Else
    Private WithEvents mImage As PictureBox
Attribute mImage.VB_VarHelpID = -1
    Private mBorder As Shape
#End If
Private m_Toolbar As Toolbar
Private m_Key     As String
Private Const MouseOverBorderColor As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private Const MouseOverBackColor   As Long = &HF3D7B3 ' = RGB(179, 215, 243)
Private Const MouseDownBorderColor As Long = &HD77800 ' = RGB(  0, 120, 215)
Private Const MouseDownBackColor   As Long = &HEBBC80 ' = RGB(128, 188, 235)
Private m_IsMouseDown As Boolean
Private m_MouseDownX As Single
Private m_MouseDownY As Single

#If VBA Then
Friend Sub New_(aImage As MSForms.Image, aKey As String)
#Else
Friend Sub New_(aImage As PictureBox, aBorder As Shape, aKey As String)
#End If
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

#If VBA Then
Private Sub mImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#Else
Private Sub mImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#End If
    If Button = EMouseButtons.MButtonLeft Then
        m_IsMouseDown = True
        m_MouseDownX = X
        m_MouseDownY = Y
        DrawButton_MouseDown
    End If
End Sub

#If VBA Then
Private Sub mImage_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#Else
Private Sub mImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
#End If
    m_Toolbar.SelectButton Me
    If m_IsMouseDown And (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
        DrawButton_MouseDown
        Exit Sub
    End If
    DrawButton_MouseOver
End Sub

#If VBA Then
Private Sub mImage_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
#Else
Private Sub mImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#End If
    If Button = EMouseButtons.MButtonLeft Then
        m_IsMouseDown = False
        Deselect
        If (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
            m_Toolbar.OnClick Me
        End If
    End If
End Sub

Private Sub DrawButton_Clear()
    mBorder.BorderStyle = BorderStyleConstants.vbTransparent 'fmBorderStyleNone
    mBorder.BackStyle = fmBackStyleTransparent
End Sub

Private Sub DrawButton_MouseOver()
    mBorder.BorderStyle = BorderStyleConstants.vbBSSolid 'fmBorderStyle.fmBorderStyleSingle
    mBorder.BorderColor = MouseOverBorderColor ' &HEBBC80 'RGB(128, 188, 235)
    mBorder.BackStyle = fmBackStyleOpaque
    mBorder.BackColor = MouseOverBackColor     ' &HF3D7B3 'RGB(179, 215, 243)
End Sub

Private Sub DrawButton_MouseDown()
    mBorder.BorderStyle = BorderStyleConstants.vbBSSolid 'fmBorderStyleSingle
    mBorder.BorderColor = MouseDownBorderColor ' &HD77800 'RGB(0, 120, 215)
    mBorder.BackStyle = Backstyleconstants.fmBackStyleOpaque
    mBorder.BackColor = MouseDownBackColor     ' &HF3D7B3 'RGB(179, 215, 243)
End Sub

