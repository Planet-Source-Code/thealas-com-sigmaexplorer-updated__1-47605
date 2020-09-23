VERSION 5.00
Begin VB.Form frm_Fullsrc 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Full Screen Preview"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic_Picture 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   0
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frm_Fullsrc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_fullscr"

Private Sub Form_Activate()
On Error GoTo e
    ' Get the pic
    With frm_Main
        Set pic_Picture.Picture = .img_Picture.Picture
        pic_Picture.Left = ScaleWidth / 2 - pic_Picture.Width / 2
        pic_Picture.Top = ScaleHeight / 2 - pic_Picture.Height / 2
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "form_activate": Resume Next
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    Select Case KeyCode
    Case vbKeyPageUp
        frm_Main.image_Previous
    Case vbKeyPageDown
        frm_Main.image_Next
    Case Else
        Unload Me
    End Select
    Form_Activate
End Sub

Private Sub pic_Picture_Click()
    Form_Click
End Sub

Private Sub pic_Picture_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub
