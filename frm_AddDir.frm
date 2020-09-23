VERSION 5.00
Begin VB.Form frm_AddDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create Directories"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "frm_AddDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   2745
      TabIndex        =   5
      Top             =   1530
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Done"
      Height          =   420
      Left            =   2745
      TabIndex        =   4
      Top             =   1035
      Width           =   1500
   End
   Begin VB.TextBox txt_Name 
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   45
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   420
      Left            =   2745
      TabIndex        =   2
      Top             =   540
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   420
      Left            =   2745
      TabIndex        =   1
      Top             =   45
      Width           =   1500
   End
   Begin VB.ListBox lst_Dirs 
      Height          =   2205
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   2625
   End
End
Attribute VB_Name = "frm_AddDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_adddir"


Private Sub Command1_Click()
On Error GoTo e
    Dim I&
    With txt_Name
        If Not .Text = "" Then
            For I = 0 To lst_Dirs.ListCount - 1
                If lst_Dirs.List(I) = .Text Then
                    MsgBox "That directory is allready listed !", vbExclamation, "Error"
                    Exit Sub
                End If
            Next I
            lst_Dirs.AddItem .Text
        End If
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "command1": Resume Next
End Sub

Private Sub Command2_Click()
On Error Resume Next
    lst_Dirs.RemoveItem lst_Dirs.ListIndex
End Sub

Private Sub Command3_Click()
On Error GoTo e
    Dim I&
    For I = 0 To lst_Dirs.ListCount - 1
        MkDir f_GetPath & "\" & lst_Dirs.List(I)
    Next I
    Select Case frm_Main.ifSelected
    Case 2
        frm_Main.dir_Right.Refresh
        frm_Main.dir_Right_Change
    Case 1
        frm_Main.dir_Left.Refresh
        frm_Main.Dir_Left_Change
    End Select
    Unload Me
Exit Sub
e:
    MsgBox "Invalid directory name or directory exists ! Check all directories for possible mistakes.", vbCritical, "Error"
    Exit Sub
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

