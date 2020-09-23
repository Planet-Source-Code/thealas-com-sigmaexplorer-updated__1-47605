VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_SystemInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Information"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frm_SystemInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmr_Refresh 
      Interval        =   2000
      Left            =   2790
      Top             =   1890
   End
   Begin VB.Frame Frame2 
      Caption         =   "Computer information"
      Height          =   780
      Left            =   3375
      TabIndex        =   22
      Top             =   45
      Width           =   3345
      Begin VB.TextBox txt_Name 
         Height          =   330
         Left            =   1605
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "Unknown"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cumputer Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   23
         Top             =   315
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drive information:  "
      Height          =   4110
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3210
      Begin VB.TextBox txt_Sectors 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "ERROR"
         Top             =   3645
         Width           =   1005
      End
      Begin VB.TextBox txt_Clusters 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "ERROR"
         Top             =   3330
         Width           =   1005
      End
      Begin VB.TextBox txt_FreeSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "ERROR"
         Top             =   3015
         Width           =   1005
      End
      Begin VB.TextBox txt_TotalSpace 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "ERROR"
         Top             =   2700
         Width           =   915
      End
      Begin VB.TextBox txt_BytesPerSector 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "ERROR"
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txt_SectorsPerCluster 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "ERROR"
         Top             =   2070
         Width           =   1005
      End
      Begin VB.TextBox txt_ID 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Unknown"
         Top             =   1755
         Width           =   1005
      End
      Begin VB.TextBox txt_FileSystem 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Unknown"
         Top             =   1440
         Width           =   1050
      End
      Begin VB.TextBox txt_DriveType 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "ERROR"
         Top             =   1125
         Width           =   915
      End
      Begin VB.TextBox txt_VolumeName 
         Height          =   330
         Left            =   1755
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   1275
      End
      Begin VB.DriveListBox drv_Drive 
         Height          =   315
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Sectors:"
         Height          =   195
         Left            =   510
         TabIndex        =   20
         Top             =   3645
         Width           =   990
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Clusters:"
         Height          =   195
         Left            =   495
         TabIndex        =   18
         Top             =   3330
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Free Space:"
         Height          =   195
         Left            =   615
         TabIndex        =   16
         Top             =   3015
         Width           =   870
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Space:"
         Height          =   195
         Left            =   570
         TabIndex        =   14
         Top             =   2700
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Bytes per sector:"
         Height          =   195
         Left            =   315
         TabIndex        =   12
         Top             =   2385
         Width           =   1185
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sectors per cluster:"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   2070
         Width           =   1365
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Serial ID:"
         Height          =   195
         Left            =   870
         TabIndex        =   8
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "File System:"
         Height          =   195
         Left            =   675
         TabIndex        =   6
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Drive Type:"
         Height          =   195
         Left            =   675
         TabIndex        =   4
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1035
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "System Status"
      Height          =   3255
      Left            =   3375
      TabIndex        =   25
      Top             =   900
      Width           =   3345
      Begin MSComctlLib.ProgressBar pb_Space 
         Height          =   285
         Left            =   135
         TabIndex        =   27
         ToolTipText     =   "Free Space"
         Top             =   495
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_PMemory 
         Height          =   285
         Left            =   135
         TabIndex        =   29
         ToolTipText     =   "Free Space"
         Top             =   1080
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_VMemory 
         Height          =   285
         Left            =   135
         TabIndex        =   31
         ToolTipText     =   "Free Space"
         Top             =   1665
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_MemoryLoad 
         Height          =   285
         Left            =   135
         TabIndex        =   33
         ToolTipText     =   "Free Space"
         Top             =   2250
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar pb_Total 
         Height          =   285
         Left            =   135
         TabIndex        =   35
         ToolTipText     =   "Free Space"
         Top             =   2835
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "System Resourses"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   2610
         Width           =   1560
      End
      Begin VB.Label lbl_MemLoad 
         AutoSize        =   -1  'True
         Caption         =   "Memory Load"
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   2025
         Width           =   960
      End
      Begin VB.Label lbl_VMem 
         AutoSize        =   -1  'True
         Caption         =   "Virtual Memory"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label lbl_PMem 
         AutoSize        =   -1  'True
         Caption         =   "Phsysical Memory"
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   855
         Width           =   1260
      End
      Begin VB.Label lbl_Drive 
         AutoSize        =   -1  'True
         Caption         =   "Drive Free Space:"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frm_SystemInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "SystemInfo"

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdLog_Click()
    If File_Ex(App.Path & "\PROGRAM.LOG") Then
        File_Open App.Path & "\PROGRAM.LOG", "Open"
    End If
End Sub

Private Sub drv_Drive_Change()
On Error GoTo e
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub



Public Sub Get_Info(Drive As String)
    Dim sInfo As SYSTEMINFO
    win_SystemInfo sInfo, Drive ' Get all the stuffs
    
    pb_Space.Max = CLng(sInfo.sysTotalSpace / 1048576)
    pb_Space.Value = CLng(sInfo.sysFreeSpace / 1048576)
    pb_PMemory.Value = sInfo.sysPsysicalMemory
    pb_VMemory.Value = sInfo.sysVirtualMemory
    pb_MemoryLoad = sInfo.sysMemoryLoad
    pb_Total.Value = (pb_PMemory.Value + pb_VMemory.Value + pb_MemoryLoad.Value + sInfo.sysPageFile) / 4
    lbl_PMem.Caption = "Phsysical Memory: " & CLng(pb_PMemory.Value) & " % Free"
    lbl_VMem.Caption = "Virtual Memory: " & CLng(pb_VMemory.Value) & " % Free"
    lbl_MemLoad.Caption = "Memory Load: " & CLng(pb_MemoryLoad.Value) & " %"
    lbl_Total.Caption = "Total: " & CLng(pb_Total.Value) & " % Free"
    txt_Name.Text = sInfo.sysComputerName
    txt_VolumeName.Text = sInfo.sysDriveName
    txt_FileSystem.Text = sInfo.sysFileSystem
    txt_ID.Text = sInfo.sysSerialID
    txt_SectorsPerCluster = sInfo.sysSectorsPerCluster
    txt_BytesPerSector.Text = sInfo.sysBytesPerSector
    txt_TotalSpace.Text = Round(sInfo.sysTotalSpace / 1048576, 1) & " MB"
    txt_FreeSpace.Text = Round(sInfo.sysFreeSpace / 1048576, 1) & " MB"
    txt_Clusters.Text = sInfo.sysTotalClusters
    txt_Sectors.Text = sInfo.sysTotalSectors
    txt_DriveType.Text = sInfo.sysDriveType
End Sub

Private Sub Form_Activate()
On Error GoTo e
    drv_Drive.Drive = frm_Main.drv_Drive.Drive
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Unload Me
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub


Private Sub tmr_Refresh_Timer()
On Error GoTo e
    ' Refresh it every 2 seconds
    Get_Info Left(drv_Drive, 2) & "\"
Exit Sub
e:
    Dim I
    I = MsgBox(Err.Description, vbCritical Or vbAbortRetryIgnore, "Error: " & Err)
    Select Case I
    Case vbAbort: Exit Sub
    Case vbIgnore: Resume Next
    Case vbRetry: Resume
    End Select
End Sub


