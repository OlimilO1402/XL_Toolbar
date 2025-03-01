VERSION 5.00
Begin VB.Form UserForm1 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   593
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox Check1 
      Caption         =   "DarkMode"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Frame FraToolbarFrame 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   4695
      Begin VB.Image ImgMouseEventListener 
         Appearance      =   0  '2D
         Height          =   495
         Left            =   120
         ToolTipText     =   "File Save"
         Top             =   240
         Width           =   495
      End
      Begin VB.Image ImgIcon 
         Height          =   345
         Left            =   180
         Picture         =   "UserForm1.frx":0000
         ToolTipText     =   "File Save"
         Top             =   300
         Width           =   375
      End
      Begin VB.Shape ShpTBButtonBorder 
         Height          =   495
         Left            =   120
         Shape           =   5  'Gerundetes Quadrat
         Top             =   240
         Width           =   495
      End
      Begin VB.Image TBBackground 
         Appearance      =   0  '2D
         Height          =   735
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   4455
      End
   End
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents mToolbar1 As Toolbar

Private MouseOverBorderColor As Long '= SystemColorConstants.vbHighlight '  &HD77800 ' = RGB(  0, 120, 215) = &HEBBC80 ' = RGB(128, 188, 235)
Private MouseOverBackColor   As Long '= &HF3D7B3 ' = RGB(179, 215, 243)
Private MouseDownBorderColor As Long '= SystemColorConstants.vbHighlight '  &HD77800 ' = RGB(  0, 120, 215)
Private MouseDownBackColor   As Long '= &HEBBC80 ' = RGB(128, 188, 235)

Private m_IsMouseDown As Boolean
Private m_MouseDownX As Single
Private m_MouseDownY As Single
Private m_IsDarkMode As Boolean

Private Const fmBackStyleTransparent As Long = 0
Private Const fmBackStyleOpaque      As Long = 1
Private Const fmBorderStyleNone      As Long = 0
Private Const fmBorderStyleSingle    As Long = 1 '0 wäre transparent

Private Sub Check1_Click()
    Me.IsDarkMode = Check1.Value = vbChecked
    
End Sub

Private Sub Form_Load()
    
    Me.ShpTBButtonBorder.Shape = ShapeConstants.vbShapeRoundedRectangle
    'Me.ShpTBButtonBorder.RoundedCornerSize = 4
    ShpTBButtonBorder.BorderStyle = 0
    
    IsDarkMode = False
End Sub

Public Property Let IsDarkMode(ByVal Value As Boolean)
    m_IsDarkMode = Value
    If m_IsDarkMode Then
        ShpTBButtonBorder.BorderStyle = 0
        Me.BackColor = RGB(23, 23, 29)
        FraToolbarFrame.BackColor = RGB(53, 52, 58)
        FraToolbarFrame.BorderStyle = 0
        MouseOverBorderColor = RGB(61, 68, 77) 'SystemColorConstants.vbHighlight
        MouseOverBackColor = RGB(74, 73, 78) 'RGB(51, 58, 69) '&HF3D7B3
        MouseDownBorderColor = RGB(46, 53, 61) 'SystemColorConstants.vbHighlight
        MouseDownBackColor = RGB(92, 92, 97) 'RGB(54, 61, 72) '&HEBBC80   '
    Else
        Me.BackColor = SystemColorConstants.vbButtonFace
        FraToolbarFrame.BackColor = SystemColorConstants.vbButtonFace
        FraToolbarFrame.BorderStyle = 1
        MouseOverBorderColor = SystemColorConstants.vbHighlight
        MouseOverBackColor = &HF3D7B3
        MouseDownBorderColor = SystemColorConstants.vbHighlight
        MouseDownBackColor = &HEBBC80   '
    End If
End Property

'Private Sub Form_Initialize()
'    'Set mToolbar1 = MNew.Toolbar(Me, FraToolbar)
'    'mToolbar1.AddButton MNew.ToolbarButton(Me.Image1, Shape1, "Save")
'    'mToolbar1.AddButton MNew.ToolbarButton(Me.Image2, Shape2, "Play")
'    'mToolbar1.AddButton MNew.ToolbarButton(Me.Image3, Shape3, "Pause")
'    'mToolbar1.AddButton MNew.ToolbarButton(Me.Image4, Shape4, "Stop")
'    'mToolbar1.DeselectAll
'    'Set mToolbar1 = MNew.Toolbar(Me, FraToolbar)
'End Sub
'
'Private Sub mToolbar1_Click(Btn As ToolbarButton)
'    Select Case Btn.Key
'    Case "Save":  ToolbarButtonSave_Click
'    Case "Play":  ToolbarButtonPlay_Click
'    Case "Pause": ToolbarButtonPause_Click
'    Case "Stop":  ToolbarButtonStop_Click
'    End Select
'End Sub
'
'Private Sub ToolbarButtonSave_Click()
'    MsgBox "Save Button clicked!"
'End Sub
'
'Private Sub ToolbarButtonPlay_Click()
'    MsgBox "Play Button clicked!"
'End Sub
'
'Private Sub ToolbarButtonPause_Click()
'    MsgBox "Pause Button clicked!"
'End Sub
'
'Private Sub ToolbarButtonStop_Click()
'    MsgBox "Stop Button clicked!"
'End Sub


'Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub
'
Private Sub FraToolbarFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Toolbar_SelectButton Nothing
End Sub
'
'Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    '
'End Sub

Public Sub Toolbar_SelectButton(aBtn As ToolbarButton)
    DrawButton_Clear
'    Dim v, obj As ToolbarButton
'    For Each v In m_Buttons
'        Set obj = v
'        If Not obj Is aBtn Then
'            obj.Deselect
'        End If
'    Next
End Sub

Friend Sub Deselect()
    DrawButton_Clear
End Sub


'Private Sub TBBackground_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    ImgMouseEventListener_MouseDown Button, Shift, X, Y
'End Sub
Private Sub TBBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'ImgMouseEventListener_MouseMove Button, Shift, X, Y
    DrawButton_Clear
End Sub
'Private Sub TBBackground_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    ImgMouseEventListener_MouseUp Button, Shift, X, Y
'End Sub

Private Sub ImgIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMouseEventListener_MouseDown Button, Shift, X, Y
End Sub
Private Sub ImgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMouseEventListener_MouseMove Button, Shift, X, Y
End Sub
Private Sub ImgIcon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgMouseEventListener_MouseUp Button, Shift, X, Y
End Sub

Private Sub ImgMouseEventListener_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        m_IsMouseDown = True
        m_MouseDownX = X
        m_MouseDownY = Y
        DrawButton_MouseDown
    End If
End Sub

Private Sub ImgMouseEventListener_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Toolbar_SelectButton Me
    If m_IsMouseDown And (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
        DrawButton_MouseDown
        Exit Sub
    End If
    DrawButton_MouseOver
End Sub

Private Sub ImgMouseEventListener_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        m_IsMouseDown = False
        Deselect
        If (Abs(m_MouseDownX - X) < 25 And Abs(m_MouseDownY - Y) < 25) Then
            'm_Toolbar.OnClick Me
        End If
    End If
End Sub

Private Sub DrawButton_Clear()
    ShpTBButtonBorder.BorderStyle = fmBorderStyleNone
    ShpTBButtonBorder.BackStyle = fmBackStyleTransparent
End Sub

Private Sub DrawButton_MouseOver()
    ShpTBButtonBorder.BorderStyle = fmBorderStyleSingle
    ShpTBButtonBorder.BorderColor = MouseOverBorderColor ' &HEBBC80 'RGB(128, 188, 235)
    ShpTBButtonBorder.BackStyle = fmBackStyleOpaque
    ShpTBButtonBorder.BackColor = MouseOverBackColor     ' &HF3D7B3 'RGB(179, 215, 243)
End Sub

Sub DrawButton_MouseDown()
    ShpTBButtonBorder.BorderStyle = fmBorderStyleSingle
    ShpTBButtonBorder.BorderColor = MouseDownBorderColor
    ShpTBButtonBorder.BackStyle = fmBackStyleOpaque
    ShpTBButtonBorder.BackColor = MouseDownBackColor
End Sub
