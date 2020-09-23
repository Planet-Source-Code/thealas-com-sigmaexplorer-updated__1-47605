VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Moving 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Moving Files..."
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   Icon            =   "frm_Moving.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   4365
      TabIndex        =   8
      Top             =   2025
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar pb_Progress 
      Height          =   240
      Left            =   90
      TabIndex        =   7
      Top             =   2205
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txt_Now 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1530
      Width           =   4155
   End
   Begin VB.TextBox txt_Dest 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   900
      Width           =   4155
   End
   Begin VB.TextBox txt_Source 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   90
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   270
      Width           =   4155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   1935
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Now moving file:"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1305
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Source:"
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Destination:"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   675
      Width           =   840
   End
End
Attribute VB_Name = "frm_Moving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_moving"

Public Cancel As Boolean

Private Sub Command1_Click()
    Cancel = True
End Sub
