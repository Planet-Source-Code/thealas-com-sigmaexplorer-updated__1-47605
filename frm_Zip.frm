VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Zip 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create  ZIP Archive - version 1.0"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frm_Zip.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lst_Files 
      Height          =   1815
      Left            =   180
      TabIndex        =   12
      Top             =   3690
      Width           =   4335
   End
   Begin VB.CheckBox ch_SavePath 
      Caption         =   "&Save path information"
      Height          =   240
      Left            =   180
      TabIndex        =   11
      Top             =   3285
      Width           =   2625
   End
   Begin VB.CheckBox ch_AddSub 
      Caption         =   "&Add sub-dirs"
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   2925
      Value           =   1  'Checked
      Width           =   2625
   End
   Begin VB.CheckBox ch_Dos 
      Caption         =   "&DOS File format"
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   2565
      Width           =   2625
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Create"
      Default         =   -1  'True
      Height          =   420
      Left            =   3285
      TabIndex        =   8
      Top             =   810
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   3285
      TabIndex        =   7
      Top             =   1305
      Width           =   1275
   End
   Begin MSComctlLib.Slider sld_Comp 
      Height          =   420
      Left            =   90
      TabIndex        =   5
      Top             =   1935
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   741
      _Version        =   393216
      SelStart        =   10
      Value           =   10
   End
   Begin VB.ComboBox cbo_Action 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frm_Zip.frx":000C
      Left            =   90
      List            =   "frm_Zip.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1125
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse"
      Height          =   420
      Left            =   3285
      TabIndex        =   2
      Top             =   315
      Width           =   1275
   End
   Begin VB.TextBox txt_File 
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Compression:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1620
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Action:"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   855
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archive Path:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "frm_Zip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_zip"

Public lFiles As New Collection

Private Sub Command1_Click()
    Dim f$
    f = frm_Main.if_ShowDirs
    If f <> "" Then
        txt_File.Text = f & "Archive.zip"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo e ' Well... I dont know what just user might do to in here :)
    Dim msg
    Me.Hide
    Zip_Add2Archive txt_File.Text, lFiles, cbo_Action.ListIndex + 1, CBool(-ch_SavePath.Value), CBool(-ch_AddSub.Value), CBool(-ch_Dos.Value), sld_Comp.Value
    Unload Me
Exit Sub
e:
    msg = MsgBox("There was an error while trying to create archive: " & Err.Description, vbCritical Or vbAbortRetryIgnore, "Err: " & Err.Number)
    Select Case msg
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
        Case vbAbort: Exit Sub
    End Select
    Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim I&, f$
    txt_File.Text = f_GetPath & "\" & "Archive.zip"
    cbo_Action.ListIndex = 0
    For I = 0 To lFiles.Count - 1
        lFiles.Remove I
    Next I
    With frm_Main.if_LW
        For I = 1 To .ListItems.Count
            If .ListItems(I).Selected = True Then
                If .ListItems(I).SubItems(1) = "<DIR>" Then
                    f = f_GetPath & "\" & .ListItems(I).Text & "\*.*"
                Else
                    f = f_GetPath & "\" & .ListItems(I).Text
                End If
                lFiles.Add f
                lst_Files.AddItem f
            End If
        Next I
    End With
End Sub

