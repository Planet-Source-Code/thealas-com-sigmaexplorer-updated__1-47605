VERSION 5.00
Begin VB.Form frm_Dirs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Directory"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "frm_Dirs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_Fav 
      Height          =   330
      Left            =   3375
      MaskColor       =   &H000000FF&
      Picture         =   "frm_Dirs.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   405
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.DriveListBox drv_Selection 
      Height          =   315
      Left            =   90
      TabIndex        =   8
      Top             =   45
      Width           =   3615
   End
   Begin VB.CheckBox ch_NoConfirmation 
      Caption         =   "&No confirmation"
      Height          =   195
      Left            =   3825
      TabIndex        =   7
      Top             =   1845
      Width           =   1455
   End
   Begin VB.CheckBox ch_Rename 
      Caption         =   "&Rename if exists"
      Height          =   195
      Left            =   3825
      TabIndex        =   6
      Top             =   1485
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&New Directory"
      Height          =   375
      Left            =   3825
      TabIndex        =   5
      Top             =   945
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3825
      TabIndex        =   4
      Top             =   495
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3825
      TabIndex        =   3
      Top             =   45
      Width           =   1500
   End
   Begin VB.TextBox txt_Add 
      Height          =   330
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   405
      Width           =   2895
   End
   Begin VB.DirListBox dir_Selection 
      Height          =   3690
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   3660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LOC:"
      Height          =   195
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   360
   End
End
Attribute VB_Name = "frm_Dirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SelectedPath As String
Const ModuleName As String = "frm_dirs"


Private Sub cmd_Fav_Click()
On Error Resume Next
    PopupMenu frm_Main.mnu_Quick
End Sub

Private Sub Command1_Click()
    SelectedPath = dir_Selection.Path
    Unload Me
End Sub

Private Sub Command2_Click()
    SelectedPath = ""
    Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo e
    Dim D$
    D = InputBox("enter directory name: ", "New Directory")
    If Not D = "" Then
        If Right(txt_Add.Text, 1) = "\" Then
            txt_Add.Text = Left(txt_Add.Text, Len(txt_Add.Text) - 1)
        End If
        MkDir txt_Add.Text & "\" & D
        dir_Selection.Path = txt_Add.Text & "\" & D
    End If
Exit Sub
e:
    MsgBox "Invalid filename !", vbExclamation, "Error": Exit Sub
End Sub

Private Sub dir_Selection_Change()
    txt_Add.Text = dir_Selection.Path
End Sub

Private Sub drv_Selection_Change()
    dir_Selection.Path = drv_Selection.Drive
End Sub

Private Sub Form_Activate()
    frm_Main.ifDirs = True
End Sub

Private Sub Form_Deactivate()
    frm_Main.ifDirs = False
End Sub

Private Sub Form_Load()
On Error GoTo e
    txt_Add.Text = f_GetPath
    dir_Selection.Path = txt_Add.Text ' Get selected path
    ch_Rename.Value = OI("chrename")
    ch_NoConfirmation.Value = OI("chcollision")
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "form_load": Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo e
    frm_Main.ifDirs = False
    SI "chrename", ch_Rename.Value
    SI "chcollision", ch_NoConfirmation.Value
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "form_unload": Resume Next
End Sub

