VERSION 5.00
Begin VB.Form UserForm1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame FraToolbar 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   3855
      Begin VB.PictureBox Image4 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1800
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   120
         Width           =   495
         Begin VB.Shape Shape4 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Image3 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   120
         Width           =   495
         Begin VB.Shape Shape3 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Image2 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   600
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   120
         Width           =   495
         Begin VB.Shape Shape2 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
      Begin VB.PictureBox Image1 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   120
         Width           =   495
         Begin VB.Shape Shape1 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mToolbar1 As Toolbar
Attribute mToolbar1.VB_VarHelpID = -1

Private Sub Form_Initialize()
    Set mToolbar1 = MNew.Toolbar(Me, FraToolbar)
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image1, Shape1, "Save")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image2, Shape2, "Play")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image3, Shape3, "Pause")
    mToolbar1.AddButton MNew.ToolbarButton(Me.Image4, Shape4, "Stop")
    mToolbar1.DeselectAll
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

