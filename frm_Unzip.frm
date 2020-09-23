VERSION 5.00
Begin VB.Form frm_Unzip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Unpack ZIP File version 1.0"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frm_Unzip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "&New Dir"
      Height          =   420
      Left            =   2835
      TabIndex        =   8
      Top             =   1125
      Width           =   1500
   End
   Begin VB.CheckBox ch_Selected 
      Caption         =   "Just &Selected"
      Height          =   240
      Left            =   2970
      TabIndex        =   7
      Top             =   2700
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   2835
      TabIndex        =   6
      Top             =   585
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Unzip"
      Default         =   -1  'True
      Height          =   420
      Left            =   2835
      TabIndex        =   5
      Top             =   45
      Width           =   1500
   End
   Begin VB.CheckBox ch_Over 
      Caption         =   "&Overwrite"
      Height          =   240
      Left            =   2970
      TabIndex        =   4
      Top             =   3060
      Width           =   1770
   End
   Begin VB.CheckBox ch_Paths 
      Caption         =   "&Use paths"
      Height          =   240
      Left            =   2970
      TabIndex        =   3
      Top             =   3420
      Value           =   1  'Checked
      Width           =   1770
   End
   Begin VB.TextBox txt_Add 
      Height          =   330
      Left            =   495
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   2220
   End
   Begin VB.DirListBox dir_Selection 
      Height          =   3240
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "LOC:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   360
   End
End
Attribute VB_Name = "frm_Unzip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_unzip"


Private Sub Command1_Click()
On Error GoTo e
    Dim zFiles As New Collection, I&
    With frm_Main.if_LW
        If ch_Selected Then
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    zFiles.Add .ListItems(I).Text ' Add the selected to the collection
                End If
            Next I
        Else
            For I = 1 To .ListItems.Count
                If Not .ListItems(I).Text = "<...>" Then zFiles.Add .ListItems(I).Text
            Next I
        End If
    End With
    Me.Hide
    ' And simply extract the collection
    Zip_Extract zFiles, zipDefault, CBool(-ch_Paths.Value), CBool(-ch_Over.Value), txt_Add.Text
    MsgBox "The files you have unziped are now in: " & vbCrLf & txt_Add.Text, vbInformation, "Unzip"
    Unload Me
Exit Sub
e:
    I = MsgBox("Zip operation failure: " & Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub

Private Sub Command2_Click()
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

Private Sub Form_Load()
    txt_Add.Text = f_GetPath
    dir_Selection.Path = txt_Add.Text
End Sub
