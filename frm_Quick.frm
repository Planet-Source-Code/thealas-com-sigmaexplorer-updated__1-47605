VERSION 5.00
Begin VB.Form frm_Quick 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Items"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frm_Quick.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   5040
      TabIndex        =   4
      Top             =   540
      Width           =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   420
      Left            =   5040
      TabIndex        =   3
      Top             =   45
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clean"
      Height          =   420
      Left            =   5040
      TabIndex        =   2
      Top             =   1530
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Remove"
      Height          =   420
      Left            =   5040
      TabIndex        =   1
      Top             =   1035
      Width           =   1320
   End
   Begin VB.ListBox lst_List 
      Height          =   2985
      ItemData        =   "frm_Quick.frx":000C
      Left            =   45
      List            =   "frm_Quick.frx":000E
      TabIndex        =   0
      Top             =   45
      Width           =   4875
   End
End
Attribute VB_Name = "frm_Quick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_quick"

Private Sub Command1_Click()
    With lst_List
        .RemoveItem .ListIndex
    End With
End Sub

Private Sub Command2_Click()
    lst_List.Clear
End Sub

Private Sub Command3_Click()
On Error GoTo e
    ' Save them
    Dim I&
    Open App.Path & "\ROCKDIRS.TXT" For Output As #1
        For I = 0 To lst_List.ListCount - 1
            Print #1, lst_List.List(I)
        Next I
    Close #1
    MsgBox "After you restart the program, the quick menu will be refreshed.", vbInformation
    Unload Me
Exit Sub
e:
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo e
    ' Load the quick items into list
    Dim I&, Inp$
    Open App.Path & "\ROCKDIRS.TXT" For Input As #1
        Do Until EOF(1)
            Input #1, Inp
            lst_List.AddItem Inp
        Loop
    Close #1
Exit Sub
e:
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub

