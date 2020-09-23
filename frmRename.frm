VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Deleteing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Deleteing Files"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "frmRename.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   3780
      TabIndex        =   3
      Top             =   900
      Width           =   1455
   End
   Begin VB.TextBox txt_Now 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3480
   End
   Begin MSComctlLib.ProgressBar pb_Progress 
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   1080
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Progress:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Deliting:"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   570
   End
End
Attribute VB_Name = "frm_Deleteing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancel As Boolean
Const ModuleName As String = "frm_adddir"


Private Sub Command1_Click()
    Cancel = True
End Sub
