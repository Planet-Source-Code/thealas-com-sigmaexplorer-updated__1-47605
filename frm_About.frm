VERSION 5.00
Begin VB.Form frm_About 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Sigma Explorer"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   45
      Picture         =   "frm_About.frx":000C
      ScaleHeight     =   3900
      ScaleWidth      =   1500
      TabIndex        =   7
      Top             =   45
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   3510
      Width           =   1005
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "www.hallsoft.tk"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2340
      MouseIcon       =   "frm_About.frx":4BC5
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   2880
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "www.freevbcode.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   2115
      MouseIcon       =   "frm_About.frx":4D17
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2340
      Width           =   1755
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "www.planetsourcecode.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   225
      Left            =   1800
      MouseIcon       =   "frm_About.frx":4E69
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2610
      Width           =   2340
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frm_About.frx":4FBB
      ForeColor       =   &H00FFC0C0&
      Height          =   1545
      Left            =   1980
      TabIndex        =   2
      Top             =   900
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Copyright (C) Hallsoft 2003"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   1890
      TabIndex        =   1
      Top             =   450
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Sigma Explorer"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   90
      Width           =   2040
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   915
      Left            =   1080
      Top             =   2295
      Width           =   4110
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Sigma Explorer " & App.Major & "." & App.Minor
End Sub

Private Sub Label6_Click()
On Error GoTo e
    File_Open "www.planetsourcecode.com", "Open"
Exit Sub
e:
    MsgBox "Error opening the link"
    Exit Sub
End Sub

Private Sub Label7_Click()
On Error GoTo e
    File_Open "www.freevbcode.com", "Open"
Exit Sub
e:
    MsgBox "Error opening the link"
    Exit Sub
End Sub

Private Sub Label8_Click()
On Error GoTo e
    File_Open "www.hallsoft.com", "Open"
Exit Sub
e:
    MsgBox "Error opening the link, try in internet explorer by you self."
    Exit Sub
End Sub
