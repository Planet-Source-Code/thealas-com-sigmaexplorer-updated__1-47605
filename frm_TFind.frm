VERSION 5.00
Begin VB.Form frm_TFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find text..."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   Icon            =   "frm_TFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   630
      Width           =   1410
   End
   Begin VB.CommandButton cmd_Find 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   135
      Width           =   1410
   End
   Begin VB.CheckBox WholeWord 
      Caption         =   "&Whole Word"
      Height          =   240
      Left            =   1395
      TabIndex        =   5
      Top             =   1035
      Width           =   1500
   End
   Begin VB.CheckBox Case 
      Caption         =   "&Match Case"
      Height          =   240
      Left            =   1395
      TabIndex        =   4
      Top             =   720
      Width           =   1500
   End
   Begin VB.OptionButton All 
      Caption         =   "&All"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   1035
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.OptionButton Down 
      Caption         =   "&Down"
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   720
      Width           =   1050
   End
   Begin VB.TextBox Text 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   270
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter String:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   870
   End
End
Attribute VB_Name = "frm_TFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_tfind"

Private Sub cmd_Cancel_Click()
    Unload Me
End Sub

Private Sub cmd_Find_Click()
On Error GoTo e
    Dim rOpt&, rStart&, rEnd&, rFind&
    ' This was in frm_main, but I've put it here
    ' and, this time there are no opt_ and ch_ shits
    With frm_Main.txt_Edit
        If cmd_Find.Caption = "&Find Next" Then
            Down.Value = True
            .SelStart = .SelStart + .SelLength
        End If
        If frm_TFind.Case.Value = 1 Then rOpt = rtfMatchCase
        If WholeWord.Value = 1 Then rOpt = rOpt Or rtfWholeWord
        If frm_TFind.Down.Value Then
            rStart = .SelStart
            rEnd = -1
        End If
        If frm_TFind.All.Value = True Then
            rStart = 1
            rEnd = -1
        End If
        rFind = .Find(Text.Text, rStart, rEnd, rOpt)
        If rFind = -1 Then MsgBox "Finished search.", vbExclamation, "Find"
    End With
    cmd_Find.Caption = "&Find Next"
    frm_Main.txt_Edit.SetFocus
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "text_find": Resume Next
End Sub

Private Sub Text_Change()
    cmd_Find.Caption = "&Find"
End Sub
