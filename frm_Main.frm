VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "AGRICHEDIT.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Sigma Explorer"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6555
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList img_Show 
      Left            =   315
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":13CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":152A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1686
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2382
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":27D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2B2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb_ProgressBar 
      Height          =   195
      Left            =   45
      TabIndex        =   12
      Top             =   4935
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar sb_StatusBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   4845
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   3969
            MinWidth        =   3969
            Text            =   "Progress"
            TextSave        =   "Progress"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2593
            MinWidth        =   2381
            Text            =   "Size 0 Kb / 0 Files  "
            TextSave        =   "Size 0 Kb / 0 Files  "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1852
            MinWidth        =   1852
            TextSave        =   "8/13/2003"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "3:33 h"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_TextEdit 
      Left            =   945
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":30B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":31CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":32E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":33FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3516
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":362E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3746
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":385E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3976
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3BAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_Title 
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   436
      TabIndex        =   22
      Top             =   765
      Width           =   6540
      Begin VB.Label lbl_Right 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Files"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3195
         TabIndex        =   24
         Top             =   0
         Width           =   3330
      End
      Begin VB.Label lbl_Left 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Directories"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   3120
      End
   End
   Begin VB.PictureBox pic_DBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   16
      Top             =   4470
      Width           =   6555
      Begin VB.CommandButton cmd_Operation 
         Caption         =   "F8-Delete"
         Height          =   330
         Index           =   4
         Left            =   4680
         TabIndex        =   21
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Operation 
         Caption         =   "F7-New Directory"
         Height          =   330
         Index           =   3
         Left            =   3510
         TabIndex        =   20
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Operation 
         Caption         =   "F6-Move"
         Height          =   330
         Index           =   2
         Left            =   2340
         TabIndex        =   19
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Operation 
         Caption         =   "F5-Copy"
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   18
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmd_Operation 
         Caption         =   "F4-Edit"
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1140
      End
   End
   Begin VB.PictureBox pic_SizeBar 
      BorderStyle     =   0  'None
      Height          =   22500
      Left            =   3120
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   7
      Top             =   840
      Width           =   60
   End
   Begin VB.PictureBox pic_IconFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   3735
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   10
      Top             =   2610
      Visible         =   0   'False
      Width           =   510
      Begin VB.PictureBox pic_Icon 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList img_ListView2 
      Left            =   1575
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3D06
            Key             =   "<DIR>"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":45FA
            Key             =   "ZIP"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":4EEE
            Key             =   "My Documents"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":57E2
            Key             =   "My Pictures"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":60D6
            Key             =   "System"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":69CA
            Key             =   "Worp"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_ListView1 
      Left            =   945
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox pic_Right 
      Height          =   3255
      Left            =   3195
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   3
      Top             =   1125
      Width           =   3345
      Begin VB.DirListBox dir_Left 
         Height          =   990
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSComctlLib.ListView lw_Right 
         Height          =   1320
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2328
         View            =   3
         Arrange         =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img_ListView2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   4101
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ext."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Modified"
            Object.Width           =   4366
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Attributes"
            Object.Width           =   3175
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img_ToolBar 
      Left            =   315
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":7496
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":75F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":9DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":C55E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":D036
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":D196
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":D2F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":D74A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_Bar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   1
      Top             =   420
      Width           =   6555
      Begin VB.CommandButton cmd_Fav 
         Height          =   330
         Left            =   6210
         MaskColor       =   &H000000FF&
         Picture         =   "frm_Main.frx":D8AA
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox txt_Add 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2100
         TabIndex        =   15
         Top             =   0
         Width           =   4080
      End
      Begin VB.DriveListBox drv_Drive 
         Height          =   315
         Left            =   0
         TabIndex        =   13
         Top             =   15
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LOC:"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   75
         Width           =   360
      End
   End
   Begin MSComctlLib.Toolbar tb_ToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "img_ToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up One Level"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Search"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Directories"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Text Edit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Multimedia"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Properties"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "System Info"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.Slider sld_Memo 
         Height          =   330
         Left            =   3150
         TabIndex        =   60
         ToolTipText     =   "Cashe Memory"
         Top             =   0
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         LargeChange     =   10
         Max             =   1000
         TickFrequency   =   100
      End
   End
   Begin VB.PictureBox pic_Show 
      Height          =   4785
      Left            =   3735
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   53
      Top             =   4860
      Visible         =   0   'False
      Width           =   3795
      Begin MSComctlLib.Toolbar tb_Show 
         Height          =   390
         Left            =   270
         TabIndex        =   56
         Top             =   3555
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "img_Show"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Export"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Full Screen Preview"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Image Properties"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Previous"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Next"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom out"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom in"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Play"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox pic_Holder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3210
         Left            =   0
         ScaleHeight     =   214
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   54
         Top             =   0
         Width           =   3750
         Begin VB.PictureBox pic_Save 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   0
            ScaleHeight     =   34
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   57
            Top             =   0
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Image img_Picture 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   990
            Top             =   900
            Width           =   1635
         End
      End
      Begin MSComDlg.CommonDialog cd_Show 
         Left            =   1305
         Top             =   4500
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Export image"
         Filter          =   "Bitmap File (*.BMP)|*.bmp"
      End
      Begin VB.Label Label9 
         Caption         =   "Use PGUP and PGDOWN for browsing in FullSrc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   58
         Top             =   4005
         Width           =   3615
      End
      Begin VB.Label lbl_Info 
         Alignment       =   2  'Center
         Caption         =   "Media Info"
         Height          =   240
         Left            =   45
         TabIndex        =   55
         Top             =   3285
         Width           =   3660
      End
   End
   Begin VB.PictureBox pic_Edit 
      Height          =   2490
      Left            =   3735
      ScaleHeight     =   162
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   363
      TabIndex        =   48
      Top             =   4860
      Visible         =   0   'False
      Width           =   5505
      Begin MSComDlg.CommonDialog cd_Edit 
         Left            =   3285
         Top             =   675
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         Filter          =   "All Files (*.*)|*.*|Rich Text (*.RTF)|*.rtf"
      End
      Begin MSComctlLib.Toolbar tb_Edit 
         Height          =   390
         Left            =   45
         TabIndex        =   52
         Top             =   0
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         ImageList       =   "img_TextEdit"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save As"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Font"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Word Wrap"
               ImageIndex      =   11
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Key Commands"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Fullscreen"
               ImageIndex      =   13
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar sb_Text 
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   2070
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   5
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   1
               AutoSize        =   2
               Enabled         =   0   'False
               Object.Width           =   1323
               MinWidth        =   1323
               TextSave        =   "CAPS"
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               AutoSize        =   2
               Object.Width           =   1323
               MinWidth        =   1323
               TextSave        =   "NUM"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   3
               AutoSize        =   2
               Object.Width           =   1323
               MinWidth        =   1323
               TextSave        =   "INS"
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox txt_Edit 
         Height          =   1725
         Left            =   -45
         TabIndex        =   49
         Top             =   360
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   3043
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frm_Main.frx":DBEC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pic_Left 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   213
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   2
      Top             =   1125
      Width           =   3120
      Begin VB.DirListBox dir_Right 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSComctlLib.ListView lw_Left 
         Height          =   1320
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2328
         View            =   3
         Arrange         =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "img_ListView1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ext."
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Modified"
            Object.Width           =   15875
         EndProperty
      End
   End
   Begin VB.PictureBox pic_Search 
      Height          =   5235
      Left            =   3735
      ScaleHeight     =   5175
      ScaleWidth      =   2880
      TabIndex        =   25
      Top             =   4860
      Visible         =   0   'False
      Width           =   2940
      Begin RichTextLib.RichTextBox txt_Path 
         Height          =   330
         Left            =   135
         TabIndex        =   61
         ToolTipText     =   "Enter where to start, use C: instead of C:\"
         Top             =   1620
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   582
         _Version        =   393217
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frm_Main.frx":DCCE
      End
      Begin VB.CommandButton cmd_GoTo 
         Caption         =   "&Go to Dir"
         Height          =   375
         Left            =   1485
         TabIndex        =   59
         Top             =   2025
         Width           =   1275
      End
      Begin VB.CheckBox ch_Fast 
         Caption         =   "&WORP SPEED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   47
         Top             =   4815
         Width           =   2625
      End
      Begin VB.TextBox txt_Size2 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         TabIndex        =   46
         Top             =   3600
         Width           =   960
      End
      Begin VB.OptionButton opt_Eq 
         Caption         =   "><"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         TabIndex        =   43
         Top             =   3285
         Width           =   645
      End
      Begin VB.CommandButton cmd_Help 
         Caption         =   "?"
         Height          =   330
         Left            =   2520
         TabIndex        =   42
         Top             =   315
         Width           =   330
      End
      Begin VB.CommandButton cmd_Search 
         Caption         =   "&Start Search"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Left            =   135
         TabIndex        =   41
         Top             =   2025
         Width           =   1275
      End
      Begin VB.CheckBox sh_Case 
         Caption         =   "&Case sensitive"
         Enabled         =   0   'False
         Height          =   240
         Left            =   135
         TabIndex        =   40
         Top             =   4455
         Width           =   2625
      End
      Begin VB.CheckBox ch_Sub 
         Caption         =   "S&earch subfolders"
         Height          =   240
         Left            =   135
         TabIndex        =   39
         Top             =   4095
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.TextBox txt_Size 
         Height          =   330
         Left            =   180
         TabIndex        =   37
         Top             =   3600
         Width           =   960
      End
      Begin VB.OptionButton opt_Smaller 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1395
         TabIndex        =   36
         Top             =   3285
         Width           =   465
      End
      Begin VB.OptionButton opt_Bigger 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   945
         TabIndex        =   35
         Top             =   3285
         Width           =   465
      End
      Begin VB.TextBox txt_Date2 
         Height          =   330
         Left            =   1710
         TabIndex        =   34
         Top             =   2790
         Width           =   1095
      End
      Begin VB.TextBox txt_Date1 
         Height          =   330
         Left            =   135
         TabIndex        =   32
         Top             =   2790
         Width           =   1050
      End
      Begin VB.CheckBox ch_Date 
         Caption         =   "&Date beetween"
         Height          =   240
         Left            =   135
         TabIndex        =   31
         Top             =   2475
         Width           =   1860
      End
      Begin VB.TextBox txt_Text 
         Enabled         =   0   'False
         Height          =   330
         Left            =   135
         TabIndex        =   29
         Top             =   945
         Width           =   2670
      End
      Begin VB.TextBox txt_Search 
         Height          =   330
         Left            =   135
         TabIndex        =   27
         Top             =   315
         Width           =   2355
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "and"
         Height          =   195
         Left            =   1215
         TabIndex        =   45
         Top             =   3645
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
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
         Left            =   180
         TabIndex        =   44
         Top             =   3285
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kb"
         Height          =   195
         Left            =   2610
         TabIndex        =   38
         Top             =   3690
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "and"
         Height          =   195
         Left            =   1305
         TabIndex        =   33
         Top             =   2835
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Start Path:"
         Height          =   195
         Left            =   135
         TabIndex        =   30
         Top             =   1395
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Containing Text:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   135
         TabIndex        =   28
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Name:"
         Height          =   195
         Left            =   135
         TabIndex        =   26
         Top             =   90
         Width           =   750
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_File_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnu_File_Edit 
         Caption         =   "&Edit                                          "
         Shortcut        =   ^E
      End
      Begin VB.Menu mnu_File_Up 
         Caption         =   "&Up One Level"
      End
      Begin VB.Menu mnu_File_Clear 
         Caption         =   "Organi&ze Quick Items "
      End
      Begin VB.Menu mnu_File_S4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Select 
         Caption         =   "&Select Filter"
      End
      Begin VB.Menu mnu_File_Reverse 
         Caption         =   "&Reverse Selection"
      End
      Begin VB.Menu mnu_File_SelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnu_File_S5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_TEdit 
         Caption         =   "&Text Edit"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnu_File_S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Move 
         Caption         =   "&Move"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnu_File_Copy 
         Caption         =   "&Copy"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_File_Rename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnu_File_Delete 
         Caption         =   "&Delete"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu_File_CreateShortcut 
         Caption         =   "Create &Shortcut"
      End
      Begin VB.Menu mnu_File_NewDir 
         Caption         =   "&New Directory"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnu_File_CreateDir 
         Caption         =   "Create &Directories"
      End
      Begin VB.Menu mnu_File_Props 
         Caption         =   "&Properties"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnu_File_Refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnu_File_S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Zip 
         Caption         =   "&ZIP Pack"
      End
      Begin VB.Menu mnu_File_UnZip 
         Caption         =   "&Unpack"
      End
      Begin VB.Menu mnu_File_UzipSelected 
         Caption         =   "Unpack &Selected"
      End
      Begin VB.Menu mnu_File_S3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Close 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnu_Explorer 
      Caption         =   "&Explorer"
   End
   Begin VB.Menu mnu_Show 
      Caption         =   "&Show"
      Begin VB.Menu mnu_Show_Files 
         Caption         =   "&Files"
         Enabled         =   0   'False
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnu_Show_Dirs 
         Caption         =   "&Directories"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnu_Show_FileSearch 
         Caption         =   "&File Search"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnu_Show_FilePrev 
         Caption         =   "&Audio/Image Preview"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnu_Show_Edit 
         Caption         =   "&Text Edit"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnu_Show_S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Show_Explorer 
         Caption         =   "&Windows Explorer"
      End
      Begin VB.Menu mnu_Show_Paint 
         Caption         =   "&MS Paint"
      End
      Begin VB.Menu mnu_Show_Word 
         Caption         =   "&MS Wordpad"
      End
      Begin VB.Menu mnu_Show_DOS 
         Caption         =   "&MS-DOS Prompt"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Show_Sys 
         Caption         =   "&System Information"
      End
      Begin VB.Menu mnu_Show_S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Show_Format 
         Caption         =   "&DOS Format drive"
      End
   End
   Begin VB.Menu mnu_Quick 
      Caption         =   "&Quick"
      Begin VB.Menu mnuFav 
         Caption         =   "&Add"
         Index           =   0
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Begin VB.Menu mnu_Help_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const ModuleName As String = "frm_main"

' For sizing the bar between two sides
Public ifSizing             As Boolean
Public ifSelected           As Integer
Public ifRelative           As Boolean  ' Sizebar X cord. is always form width / 2
Public ifFilter             As String
Public ifZip                As Boolean  ' If zip arhive is active
Public ifSearch             As Boolean  ' If search archive...
Public ifAction             As String
Public ifMaxIconCashe       As Long
Public ifText               As Boolean ' If text fullscreen is active
Public ifDirs               As Boolean ' If dirs form is active
Private Sub cmd_Fav_Click()
    PopupMenu mnu_Quick
End Sub

Private Sub cmd_GoTo_Click()
On Error Resume Next
    If ifSearch Then
        if_SearchActivation False
        if_Dir.Path = if_LW.SelectedItem.SubItems(5)
    End If
End Sub

Private Sub cmd_Help_Click()
    Dim msg$
    msg = msg & "Search help: " & vbCrLf
    msg = msg & "This is the fastest possible search code, it is faster then any other known"
    msg = msg & " programs (total commander, explorer...). Here how u use it:" & vbCrLf
    msg = msg & "If you need to find testfile.tst, you can enter the file name like this:" & vbCrLf
    msg = msg & "testfile.*  or  test*.*  or  *est*.*  or  ?estfile.*  or  *file.tst ..." & vbCrLf
    msg = msg & "These are just examples, because instead of just leaving blank space like u used"
    msg = msg & " in other programs, use DOS style, and put * ? where you need. It is better "
    msg = msg & ", because you can find just the file you wanted. For bugs report or"
    msg = msg & " questions, or any other stuffs, e-mail me !"
    MsgBox msg, vbInformation, "Help"
End Sub

Private Sub cmd_Operation_Click(Index As Integer)
    Select Case Index
    Case 0
        mnu_File_Edit_Click
    Case 1
        mnu_File_Copy_Click
    Case 2
        mnu_File_Move_Click
    Case 3
        mnu_File_NewDir_Click
    Case 4
        mnu_File_Delete_Click
    End Select
End Sub

Private Sub cmd_Search_Click()
On Error GoTo e
    Dim I&, fIcon&, fName$, fPath$, fExt$, Itm As ListItem, fSize, fAttr$, fDate$, fGAttr As VbFileAttribute, Pass As Boolean
    Dim S1!, S2!, D&, fInfo As WIN32_FIND_DATA, FTime As SYSTEMTIME
    If txt_Path.Text = "" Then MsgBox "Invalid path !", vbExclamation, "Error": Exit Sub
    ' Check if date information is valid
    If ch_Date.Value = 1 Then
        If Not IsDate(txt_Date1.Text) Then MsgBox "Invalid date !", vbExclamation, "Error": Exit Sub
        If Not IsDate(txt_Date2.Text) Then MsgBox "Invalid date !", vbExclamation, "Error": Exit Sub
    End If
    File_FindFiles txt_Path.Text & "\" & txt_Search.Text, CBool(-ch_Sub.Value) ' Use the find to list all the files in collection
    if_LW.ListItems.Clear ' Now put that files in LW
    If Not ifSearch Then if_LW.ColumnHeaders.Add , , "Path", 500 ' Add an extra column for search
    ifSearch = True
    if_SearchActivation True
    if_LW.ListItems.Add(, , "<...>", , 1).SubItems(1) = "<DIR>"
    Screen.MousePointer = vbHourglass
    For I = 1 To FindFileCol.Count
        ' Get the file infos
        fName = File_ParseName(FindFileCol(I))
        fPath = File_ParsePath(FindFileCol(I))
        fInfo = GetFindFile(File_FixPath(fPath, True) & fName)
        FileTimeToSystemTime fInfo.ftLastWriteTime, FTime
        fSize = FileLen(File_FixPath(fPath, True) & fName) / 1024
        fDate = FTime.wDay & "/" & FTime.wMonth & "/" & FTime.wYear & " " & FTime.wHour & ":" & FTime.wMinute & ":" & FTime.wSecond
        D = 0
        Do Until IsNumeric(Right(fDate, 1))
            ' This will remove all the junk from the date information
            D = D + 1
            fDate = Left(fDate, Len(fDate) - D)
        Loop
        Pass = True ' If pass is false, then that file will not be listed
        If opt_Smaller.Value Then
            S1 = txt_Size.Text
            If fSize < S1 Then Pass = True Else Pass = False
        End If
        If opt_Bigger.Value Then
            S1 = txt_Size.Text
            If fSize > S1 Then Pass = True Else Pass = False
        End If
        If opt_Eq.Value Then ' This is size between two numbers
            S2 = txt_Size2.Text
            If fSize > S1 Then
                If fSize < S2 Then
                    Pass = True
                End If
            Else
                Pass = False
            End If
        End If
        If ch_Date.Value = 1 Then ' Calculate the date differences
            If DateDiff("d", DateValue(fDate), txt_Date1.Text) < 0 Then
                If DateDiff("d", DateValue(fDate), txt_Date2.Text) > 0 Then
                    Pass = True
                End If
            Else
                Pass = False
            End If
        End If
        If Pass = True Then ' If all the conditions are met, then list this file
            fExt = UCase(Right(fName, 3))
            If ch_Fast.Value = 1 Then ' It will be a bit faster without icons
                fIcon = 6 ' 'Unknown file' icon
            Else
                fIcon = Icon_AddToImageList(File_FixPath(fPath, True) & fName, fExt, img_ListView2)
            End If
            ' Now just add the listitem, and set the subitems
            Set Itm = if_LW.ListItems.Add(, , fName, , fIcon)
            Itm.SubItems(1) = fExt
            Itm.SubItems(2) = CLng(fSize)
            If Itm.SubItems(2) = "0" Then Itm.SubItems(2) = "0." & CLng(fSize)
            Itm.SubItems(3) = fDate
            fAttr = ""
            fGAttr = fInfo.dwFileAttributes
            If (fGAttr And vbArchive) Then fAttr = fAttr & "A"
            If (fGAttr And vbHidden) Then fAttr = fAttr & ",H": Itm.Ghosted = True
            If (fGAttr And vbReadOnly) Then fAttr = fAttr & ",R": Itm.Bold = True
            If (fGAttr And vbSystem) Then fAttr = fAttr & ",S": Itm.ForeColor = vbRed
            Itm.SubItems(4) = fAttr
            Itm.SubItems(5) = fPath
        End If
    Next I
    Screen.MousePointer = vbDefault
Exit Sub
e:
    If Err = 13 Then Resume Next ' If date is invalid, the date junk is too hard
    Err_Raise Err.Number, Err.Description, ModuleName, ""
    MsgBox "An error has occured, during search: " & vbCrLf & Err.Description, vbCritical, "Error: " & Err.Number
    Screen.MousePointer = vbDefault
    Exit Sub
End Sub

Public Sub Dir_Left_Change()
    ' This side is not supported yet
End Sub

Public Sub dir_Right_Change()
On Error GoTo e
    ' Here is the file manager code, list the files and dirs into ListView
    Dim Itm As ListItem, I&, FCount&, fAttr$, fGAttr As VbFileAttribute, fName$, fSName$, fIcon&, fDir$, fPath$
    Dim fInfo As WIN32_FIND_DATA, FTime As SYSTEMTIME
    With lw_Right
        .ListItems.Clear
        If if_GetIconCashe(2) > ifMaxIconCashe Then
            Set .SmallIcons = Nothing
            For I = 7 To img_ListView2.ListImages.Count
                img_ListView2.ListImages.Remove img_ListView2.ListImages.Count
            Next I
            Set .SmallIcons = img_ListView2
        End If
        fPath = f_GetPath(2)
        If Not Len(fPath) = 2 Then ' If not root: C:\, A:\, etc.
            Set Itm = .ListItems.Add(, , "<...>", , 1) ' Add the Up dir
            Itm.SubItems(1) = "<DIR>"
        End If
        'DIRECTORIES:::
        For I = dir_Right.ListIndex + 1 To dir_Right.ListCount - 1
            fName = dir_Right.List(I)
            fInfo = GetFindFile(fName) ' Here we use api, it is not much faster, but gives better info
            FileTimeToSystemTime fInfo.ftLastWriteTime, FTime ' Get real time, instead of vba
            fSName = file_StripPath(fName)
            fIcon = Icon_AddToImageList(fName, "<DIR>", img_ListView2) ' Get the icon
            Select Case LCase(fSName) ' Check if we've got the icon
            Case "my documents"
                fIcon = 3
            Case "my pictures"
                fIcon = 4
            Case "windows"
                fIcon = 5
            End Select
            Set Itm = .ListItems.Add(, , fSName, , fIcon) ' Ready to add the first column (listitem)
            Itm.SubItems(1) = "<DIR>" ' TYPE
            Itm.SubItems(3) = FTime.wDay & "/" & FTime.wMonth & "/" & FTime.wYear & " " & FTime.wHour & ":" & FTime.wMinute & ":" & FTime.wSecond ' DATE       ' MODIFIED
            fAttr = "" ' String for holding the attributes
            fGAttr = fInfo.dwFileAttributes
            If (fGAttr And vbArchive) Then fAttr = fAttr & "A" ' ARCHIVE
            If (fGAttr And vbHidden) Then fAttr = fAttr & ",H": Itm.Ghosted = True ' HIDDEN
            If (fGAttr And vbReadOnly) Then fAttr = fAttr & ",R": Itm.Bold = True ' READ ONLY
            If (fGAttr And vbSystem) Then fAttr = fAttr & ",S": Itm.ForeColor = vbRed ' SYSTEM
            Itm.SubItems(4) = fAttr
        Next I
        
        ' This will count how many dirs&files are there
        FCount = 0
        fDir = Dir(fPath & "\", vbHidden Or vbSystem) ' Include hidden and system stuffs
        While fDir <> ""
            FCount = FCount + 1
            fDir = Dir
        Wend
        
        pb_ProgressBar.Max = FCount + 1
        pb_ProgressBar.Value = 0
        ' FILES:::
        ' NO FileSystem, just api !!!
        fDir = Dir(fPath & "\" & ifFilter, vbHidden Or vbSystem) ' This will get the first file
        While fDir <> ""
            'fDir is just name of the file (and extension)
            fName = fPath & "\" & fDir
            fIcon = Icon_AddToImageList(fName, UCase(Right(fDir, 3)), img_ListView2) ' Get the icon
            fInfo = GetFindFile(fName) ' Here we use api, it is not much faster, but gives better info
            FileTimeToSystemTime fInfo.ftLastWriteTime, FTime ' Get real time, instead of vba
            Set Itm = .ListItems.Add(, , fDir, , fIcon) ' FILE NAME
            pb_ProgressBar.Value = pb_ProgressBar.Value + 1 ' Display progress
            Itm.SubItems(1) = UCase(Right(fDir, 3)) ' TYPE
            Itm.SubItems(2) = CLng(FileLen(fName) / 1024) ' SIZE
            If Itm.SubItems(2) = "0" Then Itm.SubItems(2) = "0." & CLng(FileLen(fName)) ' If smaller then 1 Kb
            Itm.SubItems(3) = FTime.wDay & "/" & FTime.wMonth & "/" & FTime.wYear & " " & FTime.wHour & ":" & FTime.wMinute & ":" & FTime.wSecond ' DATE
            fAttr = "" ' Just string for holding attributes info ('a,h,r,s')
            fGAttr = fInfo.dwFileAttributes ' get attributes, faster then fileattr
            ' Use AND, if the wanted attribute is just one of them, simple...
            If (fGAttr And vbArchive) Then fAttr = fAttr & "A"
            If (fGAttr And vbHidden) Then fAttr = fAttr & ",H": Itm.Ghosted = True
            If (fGAttr And vbReadOnly) Then fAttr = fAttr & ",R": Itm.Bold = True
            If (fGAttr And vbSystem) Then fAttr = fAttr & ",S": Itm.ForeColor = vbRed
            Itm.SubItems(4) = fAttr ' ATTRIBUTES
            ' Just extension check
            Select Case UCase(Right(fDir, 3))
            Case "BMP", "DIB", "GIF", "JPG", "WMF", "EMF", "ICO", "CUR", "WAV"
                Itm.ForeColor = vbBlue
            End Select
            fDir = Dir ' And then goes the second file, until fDir = "", what means
            'that there are no files left.
        Wend
        pb_ProgressBar.Value = 0
        ' Display the situation in status bar
        sb_StatusBar.Panels(2).Text = "Files: " & FCount & " Dirs: " & dir_Right.ListCount & " / " & File_GetLWSizeTotal(lw_Right, 2) & " Kb Total, " & " Cashe: " & if_GetIconCashe(2) & " Kb   "
        .ListItems(1).Selected = False
    End With
    txt_Add.Text = dir_Right.Path
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "dir_right_change": Resume Next
End Sub

Private Sub drv_Drive_Change()
On Error GoTo e
    if_Dir.Path = drv_Drive.Drive
Exit Sub
e:
    MsgBox Err.Description, vbExclamation, "Error: " & Err
    drv_Drive.Drive = "c:\"
    Exit Sub
End Sub

Private Sub Form_Load()
    ' NEVER put large code in here, because the smallest bug may prevent
    'program from running !
    if_LoadInfo ' Set the needed parameters, and load some from INI file
End Sub

Private Sub Form_Resize()
    if_Size
End Sub


Private Sub Form_Unload(Cancel As Integer)
    if_SaveInfo ' Save some stuffs to INI file
End Sub



Private Sub lw_Left_BeforeLabelEdit(Cancel As Integer)
    ' This side is not supported yet
End Sub

Private Sub lw_Right_AfterLabelEdit(Cancel As Integer, NewString As String)
On Error GoTo e
    ' Rename the folder or file, if user caused this event
    Name f_GetPath(2) & "\" & lw_Right.SelectedItem.Text As f_GetPath & "\" & NewString
Exit Sub
e:
    Cancel = 1
    MsgBox "Invalid directory name, or directory exists !", vbCritical, "Error"
    Err_Raise Err.Number, Err.Description, ModuleName, "lw_right_afterlabeledit": Resume Next
End Sub

Private Sub lw_Right_BeforeLabelEdit(Cancel As Integer)
    ' Prevents the user from getting into truble
    If lw_Right.SelectedItem.Text = "<...>" Then Cancel = 1
End Sub

Private Sub lw_Right_Click()
On Error Resume Next
    ' Deselect if possible
    If lw_Right.SelectedItem.Text = "<...>" Then lw_Right.SelectedItem.Selected = False
End Sub

Private Sub lw_Right_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
    Dim I&, C$
    With ColumnHeader
        ' This will put all the dirs at the top
        For I = 1 To if_LW.ListItems.Count
            If if_LW.ListItems(I).SubItems(1) = "<DIR>" Then
                If .Index = 1 Then
                    if_LW.ListItems(I).Tag = if_LW.ListItems(I).Text
                    if_LW.ListItems(I).Text = ""
                Else
                    if_LW.ListItems(I).Tag = if_LW.ListItems(I).SubItems(.Index - 1)
                    if_LW.ListItems(I).SubItems(.Index - 1) = ""
                End If
            Else
                Select Case .Index
                Case 3
                    if_LW.ListItems(I).Tag = if_LW.ListItems(I).SubItems(.Index - 1)
                    if_LW.ListItems(I).SubItems(.Index - 1) = Zeros(CLng(if_LW.ListItems(I).SubItems(.Index - 1)))
                Case 4
                    if_LW.ListItems(I).Tag = if_LW.ListItems(I).SubItems(.Index - 1)
                    if_LW.ListItems(I).SubItems(.Index - 1) = Zeros(DateDiff("d", "1/1/1980", if_LW.ListItems(I).SubItems(.Index - 1)))
                End Select
            End If
        Next I
        ' This restores the lw to old, but leaving the items order
        if_LW.SortKey = .Index - 1
        if_LW.Sorted = True
        For I = 1 To if_LW.ListItems.Count
            If if_LW.ListItems(I).Tag <> "" Then
                If .Index = 1 Then
                    if_LW.ListItems(I).Text = if_LW.ListItems(I).Tag
                    if_LW.ListItems(I).Tag = ""
                Else
                    if_LW.ListItems(I).SubItems(.Index - 1) = if_LW.ListItems(I).Tag
                    if_LW.ListItems(I).Tag = ""
                End If
            End If
        Next I
        if_LW.Sorted = False
    End With
End Sub

Private Sub lw_Right_DblClick()
On Error GoTo e ' Who knows what just might happend in here !!!
    ' Opens the selected file, or dir
    Dim fPath$, zFiles As New Collection, I&
    With lw_Right
        fPath = f_GetPath(2)
        ' DIRECTORIES
        If .SelectedItem.SubItems(1) = "<DIR>" Then
            If .SelectedItem.Text = "<...>" Then ' If Level Up
                If ifZip Then ' If we are in ZIP, close it
                    if_ZipActivation False
                    dir_Right.Refresh
                    dir_Right_Change
                End If
                If ifSearch Then ' If search results, close them
                    if_SearchActivation False
                    ifSearch = False: txt_Add.Enabled = True
                    dir_Right.Refresh
                    dir_Right_Change
                    if_SetMode 2, 1
                Else
                    dir_Right.Path = dir_Right.Path & "\.." ' Goes one level up
                End If
            Else
                dir_Right.Path = fPath & "\" & .SelectedItem.Text ' Opens the directory
            End If
        ' FILES
        Else
            If UCase(Right(.SelectedItem.Text, 3)) = "ZIP" Then ' If zip archive
                If ifSearch Then
                    MsgBox "You are in search mode, zip files cannot be opened now.", vbInformation
                Else
                    Zip_Open
                End If
            End If
            If ifZip Then ' If the ZIP archive is opened
                zFiles.Add if_LW.SelectedItem.Text
                Zip_Extract zFiles, zipDefault, False, True, win_GetDirectory & "\TEMP"
                File_Open win_GetDirectory & "\TEMP\" & if_LW.SelectedItem.Text, ifAction
            End If
            If ifSearch Then ' If search results
                File_Open .SelectedItem.SubItems(5) & .SelectedItem.Text, ifAction
            Else
                File_Open fPath & "\" & .SelectedItem.Text, ifAction ' If simple file
            End If
        End If
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "lw_right_dblclick": Resume Next
End Sub


Private Sub lw_Right_GotFocus()
    if_Selection 2 ' Select the RIGHT mode
End Sub

Public Sub if_Selection(ifSide As Integer)
    Select Case ifSide
    Case 2
        ifSelected = 2
        txt_Add.Text = dir_Right.Path
    Case 1
        ifSelected = 1
        txt_Add.Text = dir_Left.Path
    End Select
End Sub

Private Sub lw_Right_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo e
    ' This events is used for TextEdit, Wav play and Image Show
    Dim fPath$, X&, Y&, C&, Ext$
    tb_Show.Buttons("Play").Enabled = False ' Disable the play for default
    If Item.Text = "<...>" Then ' The user is prevented from selecting this file
        Item.Selected = False
        Exit Sub
    End If
    If ifSearch Then
        ' If search results are active, then select them
        fPath = Item.SubItems(5) & Item.Text
    Else
        fPath = f_GetPath(2) & "\" & Item.Text
    End If
    If (GetAttr(fPath) And vbDirectory) = False Then ' If not directory
        If pic_Show.Visible Then
            Ext = UCase(Right(Item.Text, 3)) ' Extension
            Select Case Ext
            Case "BMP", "DIB", "GIF", "JPG", "WMF", "EMF", "ICO", "CUR"
                With img_Picture
                    .Stretch = False
                    .Picture = LoadPicture(fPath) ' Load the pic
                    image_GetInfo fPath, X, Y, C ' Get the size, and color depth
                    lbl_Info.Caption = "Width: " & X & " Height: " & Y & " Depth: " & C
                    ' If ICO, then different story
                    If Ext = "ICO" Then lbl_Info.Caption = "Width: " & .Width - 2 & " Height: " & .Height - 2 & " Depth: ?"
                    ' Fit the image, if it is larger then its parent
                    If .Width > pic_Holder.Width Then
                        .Stretch = True
                        .Width = pic_Holder.Width
                    End If
                    If .Height > pic_Holder.Height Then
                        .Stretch = True
                        .Height = pic_Holder.Height
                    End If
                    ' Center it
                    .Left = pic_Holder.Width / 2 - .Width / 2
                    .Top = pic_Holder.Height / 2 - .Height / 2
                End With
            Case "WAV"
                'If wav, then.... play it.
                lbl_Info.Caption = UCase(File_ParseName(fPath)) & " - No info"
                tb_Show.Buttons("Play").Enabled = True
                sndPlaySound fPath, SND_ASYNC
            End Select
        End If
        ' Load the file if TextEdit is active
        If pic_Edit.Visible Then
            txt_Edit.LoadFile fPath
        End If
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "lw_right_itemclick: " & Item.Text: Resume Next
End Sub

Private Sub lw_Right_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lw_Right_DblClick ' Enter is same as dblclick
End Sub


Private Sub lw_Right_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo e
    ' When user relases the mouse, then show the info for all files, no need for props form.
    Dim FS!, FC&
    FS = File_GetLWSize(lw_Right, 2, FC) ' Get the size of selected files
    sb_StatusBar.Panels(3).Text = "Size: " & FS & " Kb" & " / " & FC & " Files  "
    If Button = 2 Then PopupMenu mnu_File
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "lw_right_mouseup": Resume Next
End Sub

Private Sub mnu_Explorer_Click()
On Error Resume Next
    If Not win_Function_Exist("cfexpmnu.dll", "DoExplorerMenu") Then MsgBox "Explorer plug-in is missing, contact the author about this problem.", vbCritical, "Error": Exit Sub
    ' Show the explorer menu, if u use winzip, or some other shit for windows.
    DoExplorerMenu Me.hWnd, f_GetPath & "\" & if_LW.SelectedItem.Text, 32, 0
End Sub

Private Sub mnu_File_Clear_Click()
    frm_Quick.Show vbModal, Me
End Sub

Private Sub mnu_File_Close_Click()
    Unload Me
End Sub

Private Sub mnu_File_Copy_Click()
On Error GoTo e
    Dim fDest$, fSrc$, I&, C&
    fDest = if_ShowDirs ' Browse dialog
    With if_LW
        If fDest <> "" Then
            frm_Moving.Show vbModeless, Me ' Moving form
            frm_Moving.Caption = "Copying files..." ' Change the title
            DoEvents ' Allow the user for doing something else
            ' Get the file count, for progressbar
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    C = C + 1
                End If
            Next I
            frm_Moving.pb_Progress.Max = C + 1 ' Ya see
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    ' If search results, then copy them
                    If ifSearch Then
                        fSrc = .ListItems(I).SubItems(5) & .ListItems(I).Text
                    Else
                        fSrc = f_GetPath & "\" & .ListItems(I).Text
                    End If
                    File_Copy fSrc, fDest ' Simple file/dir copy from explorer
                    With frm_Moving
                        DoEvents
                        If .Cancel Then ' If the user canceled, then stop
                            .Cancel = False
                            Unload frm_Moving
                            Exit Sub
                        End If
                        .txt_Dest.Text = fDest
                        .txt_Source.Text = fSrc
                        .txt_Now.Text = File_ParseName(fSrc)
                        .pb_Progress.Value = .pb_Progress.Value + 1
                    End With
                End If
            Next I
        End If
        frm_Moving.pb_Progress.Value = 0
        Unload frm_Moving
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_copy: " & fSrc & " TO " & fDest: Resume Next
End Sub

Private Sub mnu_File_CreateDir_Click()
    frm_AddDir.Show vbModeless, Me
End Sub

Private Sub mnu_File_CreateShortcut_Click()
On Error Resume Next
    ' I have noticed that on some computers, this stuff sometimes works, sometimes not, strange API indeed ...
    ' If u know how to realy use this api, tell me.
    Dim lFile$
    lFile = if_LW.SelectedItem.Text
    File_CreateLink f_GetPath & "\" & lFile, f_GetPath, "", Left(lFile, Len(lFile) - 4), False, ""
    if_Refresh
End Sub

Private Sub mnu_File_Delete_Click()
On Error GoTo e
    ' All the comments are in mnu_file_copy_click
    Dim fSrc$, I&, C&, msg$, zFiles As New Collection
    msg = MsgBox("Are you sure you want to delete this files ?", vbQuestion Or vbYesNo, "Delete Files")
    With if_LW
        If msg = vbYes Then
            If Not ifZip Then
                frm_Deleteing.Show vbModeless, Me
                DoEvents
                For I = 1 To .ListItems.Count
                    If .ListItems(I).Selected = True Then
                        C = C + 1
                    End If
                Next I
                frm_Deleteing.pb_Progress.Max = C + 1
                For I = 1 To .ListItems.Count
                    If .ListItems(I).Selected = True Then
                    If ifSearch Then
                        fSrc = .ListItems(I).SubItems(5) & .ListItems(I).Text
                    Else
                        fSrc = f_GetPath & "\" & .ListItems(I).Text
                    End If
                        File_Delete fSrc
                        With frm_Deleteing
                            DoEvents
                            If .Cancel Then
                                .Cancel = False
                                mnu_File_Refresh_Click
                                Unload frm_Deleteing
                                Exit Sub
                            End If
                            .txt_Now.Text = File_ParseName(fSrc)
                            .pb_Progress.Value = .pb_Progress.Value + 1
                        End With
                    End If
                Next I
            Else
                For I = 1 To .ListItems.Count
                    If .ListItems(I).Selected = True Then
                        zFiles.Add .ListItems(I).Text
                    End If
                Next I
                Zip_DeleteFiles zFiles
            End If
' Now remove all the selected items
clean:
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    .ListItems.Remove I
                    GoTo clean
                End If
            Next I
        End If
        frm_Deleteing.pb_Progress.Value = 0
        Unload frm_Deleteing
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_delete: " & fSrc: Resume Next
End Sub

Private Sub mnu_File_Edit_Click()
    ' Just change the action
    ifAction = "Edit"
    Select Case ifSelected
    Case 1
        'lw_Left_DblClick
    Case 2
        lw_Right_DblClick
    End Select
    ifAction = "Open"
End Sub

Private Sub mnu_File_Move_Click()
On Error GoTo e
    ' All the comments are in mnu_file_copy_click
    Dim fDest$, fSrc$, I&, C&
    fDest = if_ShowDirs
    With if_LW
        If fDest <> "" Then
            frm_Moving.Show vbModeless, Me
            frm_Moving.Caption = "Moving files..."
            DoEvents
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    C = C + 1
                End If
            Next I
            frm_Moving.pb_Progress.Max = C + 1
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    If ifSearch Then
                        fSrc = .ListItems(I).SubItems(5) & .ListItems(I).Text
                    Else
                        fSrc = f_GetPath & "\" & .ListItems(I).Text '!!!!!!
                    End If
                    File_Move fSrc, fDest
                    With frm_Moving
                        DoEvents
                        If .Cancel Then
                            .Cancel = False
                            Unload frm_Moving
                            Exit Sub
                        End If
                        .txt_Dest.Text = fDest
                        .txt_Source.Text = fSrc
                        .txt_Now.Text = File_ParseName(fSrc)
                        .pb_Progress.Value = .pb_Progress.Value + 1
                    End With
                End If
            Next I
clean:
            For I = 1 To .ListItems.Count
                If .ListItems(I).Selected = True Then
                    .ListItems.Remove I
                    GoTo clean
                End If
            Next I
        End If
        frm_Moving.pb_Progress.Value = 0
        Unload frm_Moving
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_move: " & fSrc & " TO " & fDest: Resume Next
End Sub

Private Sub mnu_File_NewDir_Click()
On Error GoTo e
    Dim D$, Itm As ListItem, fName$, fAttr$, I&, C&
    With if_LW
        C = 1
        ' This code will see if there are some New Directories, not something.
        For I = 1 To .ListItems.Count
            If InStr(1, .ListItems(I).Text, "New Directory ") > 0 Then ' Count them
                C = Mid(.ListItems(I).Text, 14, 999) + 1 ' And get the index
            End If
        Next I
        D = "New Directory " & C
        ' If there are too much new dirs, then mkdir may fail, so you should better name them
        MkDir f_GetPath & "\" & D
        fName = f_GetPath & "\" & D
        ' Add the dir to LW
        Set Itm = .ListItems.Add(, , D, , 1)
        Itm.SubItems(1) = "<DIR>"
        Itm.SubItems(3) = FileDateTime(fName)
        fAttr = ""
        If (GetAttr(fName) And vbArchive) Then fAttr = fAttr & "A"
        If (GetAttr(fName) And vbHidden) Then fAttr = fAttr & ",H": Itm.Ghosted = True
        If (GetAttr(fName) And vbReadOnly) Then fAttr = fAttr & ",R": Itm.Bold = True
        If (GetAttr(fName) And vbSystem) Then fAttr = fAttr & ",S": Itm.ForeColor = vbRed
        Itm.SubItems(4) = fAttr
        Itm.Selected = True
    End With
    If ifSelected = 1 Then
        lw_Left.SetFocus
        lw_Left.StartLabelEdit
        dir_Left.Refresh
    Else
        lw_Right.SetFocus
        lw_Right.StartLabelEdit
        dir_Right.Refresh
    End If
Exit Sub
e:
    MsgBox "Could not create directory: " & Err.Description, vbExclamation, "Error " & Err.Number
    Exit Sub
End Sub

Private Sub mnu_File_Open_Click()
    Select Case ifSelected
    Case 1
        'lw_Left_DblClick
    Case 2
        lw_Right_DblClick
    End Select
End Sub

Private Sub mnu_File_Props_Click()
On Error GoTo e
    ' All the props are showed in the statusbar, for multiple files. For a single file
    ' explorer properties window will do just fine !
    If ifSearch Then
        File_ShowProps if_LW.SelectedItem.SubItems(5) & if_LW.SelectedItem.Text, Me.hWnd
    Else
        File_ShowProps f_GetPath & "\" & if_LW.SelectedItem.Text, Me.hWnd
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_props": Resume Next
End Sub

Private Sub mnu_File_Refresh_Click()
    if_Refresh
End Sub

Private Sub mnu_File_Rename_Click()
On Error GoTo e
    if_LW.StartLabelEdit
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_rename": Resume Next
End Sub

Private Sub mnu_File_Reverse_Click()
    ' This will reverse the selection
    Dim I&
    For I = 1 To if_LW.ListItems.Count
        If if_LW.ListItems(I).Selected Then
            if_LW.ListItems(I).Selected = False
        Else
            if_LW.ListItems(I).Selected = True
        End If
    Next I
End Sub

Private Sub mnu_File_Select_Click()
    ' This will show only the selected (filtered) files
    Dim Filter$
    Filter = InputBox("Enter filter (*.* for all files) :", "Select Filter")
    If Not Filter = "" Then
        If Left(Filter, 2) = "*." Then
            ifFilter = Filter
        End If
        if_Refresh
    End If
End Sub

Private Sub mnu_File_SelectAll_Click()
    Dim I&
    For I = 1 To if_LW.ListItems.Count
        if_LW.ListItems(I).Selected = True
    Next I
    If ifSelected = 1 Then
        'lw_Left_MouseUp 1, 0, 1, 1
    Else
        lw_Right_MouseUp 1, 0, 1, 1
    End If
End Sub


Private Sub mnu_File_TEdit_Click()
    ' Show text editor
    If lw_Right.Visible = True Then
        if_SetMode 4, 1
    Else
        if_SetMode 1, 2
        if_SetMode 4, 1
    End If
End Sub

Private Sub mnu_File_UnZip_Click()
    If ifZip Then frm_Unzip.Show vbModal, Me
End Sub

Private Sub mnu_File_Up_Click()
On Error GoTo e
    ' Goes up one level
    If Not Len(f_GetPath) = 2 Then
        if_LW.ListItems(1).Selected = True
        Select Case ifSelected
        Case 1
            'lw_Left_DblClick
        Case 2
            lw_Right_DblClick
        End Select
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_file_up": Resume Next
End Sub

Private Sub mnu_File_UzipSelected_Click()
    If ifZip Then
        frm_Unzip.ch_Selected.Value = 1
        frm_Unzip.Show vbModal, Me
    End If
End Sub

Private Sub mnu_File_Zip_Click()
    frm_Zip.Show vbModeless, Me
End Sub

Private Sub mnu_Help_About_Click()
    frm_About.Show vbModal, Me
End Sub

Private Sub mnu_Show_Dirs_Click()
    if_SetMode 2, 1
End Sub

Private Sub mnu_Show_DOS_Click()
'    Dim dCmd$, dFile$, dSDir$, dLen&
'    dSDir = Space(512)
'    dLen = GetShortPathName(f_GetPath, dSDir, Len(dSDir))
'    dSDir = Left(dSDir, dLen)
'
'    dFile = App.Path & "\COMMAND.BAT"
'    dCmd = "CD " & dSDir
'    Open dFile For Output As #1
'        Print #1, dCmd
'    Close #1
'    Shell "CD C:\GAMES"
End Sub

Private Sub mnu_Show_Edit_Click()
    If lw_Right.Visible = True Then
        if_SetMode 4, 1
    Else
        if_SetMode 1, 2
        if_SetMode 4, 1
    End If
End Sub

Private Sub mnu_Show_Explorer_Click()
On Error Resume Next
    File_Open f_GetPath, "Open" ' Opens a path
End Sub

Private Sub mnu_Show_FilePrev_Click()
    If lw_Right.Visible = True Then
        if_SetMode 5, 1
    Else
        if_SetMode 1, 2
        if_SetMode 5, 1
    End If
End Sub


Private Sub mnu_Show_FileSearch_Click()
    If lw_Right.Visible = True Then
        if_SetMode 3, 1
    Else
        if_SetMode 1, 2
        if_SetMode 3, 1
    End If
    if_Size
End Sub

Private Sub mnu_Show_Format_Click()
    Dim sFile$, sDrive$
    sFile = App.Path & "\DOSB.BAT"
    sDrive = Left(drv_Drive.Drive, 2)
    Open sFile For Output As #1
        Print #1, "FORMAT " & sDrive
    Close #1
    Shell sFile, vbMaximizedFocus
End Sub

Private Sub mnu_Show_Paint_Click()
On Error GoTo e
    Dim pFile$, L&
    pFile = Space(512)
    L = GetShortPathName(f_GetPath & "\" & if_LW.SelectedItem.Text, pFile, Len(pFile))
    pFile = Left(pFile, L)
    Shell "pbrush " & pFile, vbNormalFocus
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_show_paint: " & pFile: Resume Next
End Sub

Private Sub mnu_Show_Sys_Click()
    ' System info
    frm_SystemInfo.Show vbModal, Me
End Sub

Private Sub mnu_Show_Word_Click()
On Error GoTo e
    Dim pFile$, L&
    pFile = Space(512)
    L = GetShortPathName(f_GetPath & "\" & if_LW.SelectedItem.Text, pFile, Len(pFile))
    pFile = Left(pFile, L)
    Shell "write " & pFile, vbNormalFocus
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnu_show_word: " & pFile: Resume Next
End Sub

Private Sub mnuFav_Click(Index As Integer)
On Error GoTo e
    ' Quick items, add or select
    If Index = 0 Then 'Add
        Open App.Path & "\ROCKDIRS.TXT" For Append As #1
            Print #1, txt_Add.Text
            Load mnuFav(mnuFav.Count)
            mnuFav(mnuFav.Count - 1).Visible = True
            mnuFav(mnuFav.Count - 1).Caption = txt_Add.Text
        Close #1
    Else ' Select
        If ifDirs Then
            frm_Dirs.dir_Selection.Path = mnuFav(Index).Caption
        Else
            txt_Add.Text = mnuFav(Index).Caption
            txt_Add_KeyPress 13
        End If
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "mnufav": Resume Next
End Sub

Private Sub opt_Bigger_Click()
    txt_Size2.Enabled = opt_Eq.Value
    txt_Size.Text = "0"
End Sub

Private Sub opt_Eq_Click()
    txt_Size2.Enabled = opt_Eq.Value
    txt_Size2.Text = "0"
End Sub

Private Sub opt_Smaller_Click()
    txt_Size2.Enabled = opt_Eq.Value
    txt_Size.Text = "0"
End Sub


Private Sub pic_DBar_Resize()
On Error Resume Next
    ' Resize the buttons
    Dim I&
    For I = 0 To cmd_Operation.Count - 1
        cmd_Operation(I).Width = pic_DBar.Width / cmd_Operation.Count - 1
        cmd_Operation(I).Left = cmd_Operation(I).Width * I
    Next I
End Sub







Private Sub pic_SizeBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    ' Activate the sizing
    If ifSearch = False Then
        If pic_Edit.Visible = False Then
            If pic_Show.Visible = False Then
                ifSizing = True: pic_SizeBar.BackColor = vbButtonShadow
            End If
        End If
    End If
End Sub

Private Sub pic_SizeBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    ' Resize pictureboxes using Mouse X
    Dim Pos&
    With pic_SizeBar
        If ifSizing Then
            Pos = .Left + X
            If Pos < 100 Then
                Pos = 100
            ElseIf Pos > Me.ScaleWidth - 100 Then
                Pos = Me.ScaleWidth - 100
            Else
                .Left = Pos ' Move the bar
            End If
        End If
    End With
End Sub

Private Sub pic_SizeBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    ifSizing = False: pic_SizeBar.BackColor = vbButtonFace
    if_Size ' Resize the stuffs
    if_Size ' Few times, to be sure ;)
End Sub

Public Sub if_SaveInfo()
On Error GoTo e
    SI "sbar", pic_SizeBar.Left
    SI "cashe", CStr(ifMaxIconCashe)
    SI "path", f_GetPath & "\"
    If Me.WindowState <> vbMaximized Then
        SI "sx", Me.Width
        SI "sy", Me.Height
    End If
    ' Clear the memory
    if_LW.ListItems.Clear
    Set if_LW.SmallIcons = Nothing
    img_ListView2.ListImages.Clear
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_saveinfo": Resume Next
End Sub

Public Sub if_LoadInfo()
On Error GoTo e
    Dim I&, fLine$
    Me.Width = OI("sx", 10245)
    Me.Height = OI("sy", 7095)
    pic_SizeBar.Left = OI("sbar", 196)
    ifMaxIconCashe = OI("cashe", 200)
    txt_Edit.RightMargin = 2500
    sld_Memo.Value = ifMaxIconCashe
    ' If sbar is set to the middle, or sized by the user
    If OI("sbarrelative", 0) = 1 Then pic_SizeBar.Left = Me.ScaleWidth / 2: ifRelative = True
    dir_Left.Path = "C:\"
    Dir_Left_Change
    dir_Right.Path = OI("path", "C:\")
    dir_Right_Change
    ' Modes for left and right sides
    if_SetMode OI("left", 2), 1
    if_SetMode OI("right", 1), 2
    ifFilter = "*.*" ' Default filter
    ifAction = "Open" ' Default action
    ifZip = False
    ' Icons Cashe in Kb
    ifMaxIconCashe = 200
    ' Set the tooltips
    For I = 1 To tb_Edit.Buttons.Count
        tb_Edit.Buttons(I).ToolTipText = tb_Edit.Buttons(I).Key
    Next I
    For I = 1 To tb_ToolBar.Buttons.Count
        tb_ToolBar.Buttons(I).ToolTipText = tb_ToolBar.Buttons(I).Key
    Next I
    For I = 1 To tb_Show.Buttons.Count
        tb_Show.Buttons(I).ToolTipText = tb_Show.Buttons(I).Key
    Next I
    For I = 1 To tb_ToolBar.Buttons.Count
        tb_ToolBar.Buttons(I).ToolTipText = tb_ToolBar.Buttons(I).Key
    Next I
    ' Open the quick items
    Open App.Path & "\ROCKDIRS.TXT" For Input As #1
        Do Until EOF(1)
            Input #1, fLine
            Load mnuFav(mnuFav.Count)
            mnuFav(mnuFav.Count - 1).Visible = True
            mnuFav(mnuFav.Count - 1).Caption = fLine
        Loop
    Close #1
    
    if_Size
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_loadinfo": Resume Next
End Sub

Public Sub if_SetMode(Choice As Long, Side As Integer)
On Error GoTo e
    ' This will set selected mode to the selected side
    Select Case Choice
    Case 1 ' FILES
        If Side = 1 Then
            lw_Left.Visible = True
            dir_Right.Visible = False
            pic_Show.Visible = False
            pic_Search.Visible = False
            pic_Edit.Visible = False
            ifSelected = 1
            lbl_Left.Caption = "Files"
        Else
            lw_Right.Visible = True
            dir_Left.Visible = False
            ifSelected = 2
            lbl_Right.Caption = "Files"
        End If
    Case 2 ' DIRS
        If ifSearch Then
            MsgBox "You must exit search results first, by going up one level", vbExclamation
        Else
            If Side = 1 Then
                lw_Left.Visible = False
                dir_Right.Visible = True
                pic_Search.Visible = False
                pic_Show.Visible = False
                pic_Edit.Visible = False
                pic_Left.BorderStyle = 0
                lbl_Left.Caption = "Directories"
            End If
        End If
    Case 3 ' SEARCH
        If Side = 1 Then
            lw_Left.Visible = False
            pic_Show.Visible = False
            dir_Right.Visible = False
            pic_Edit.Visible = False
            pic_Left.BorderStyle = 0
            pic_Search.Visible = True
            Set pic_Search.Container = pic_Left
            lbl_Left.Caption = "File Search"
        End If
        txt_Path.Text = f_GetPath
        txt_Path.SetFocus
        if_Size
    Case 4 ' TEXT EDITOR
        If Side = 1 Then
            lw_Left.Visible = False
            dir_Right.Visible = False
            pic_Show.Visible = False
            pic_Search.Visible = False
            pic_Left.BorderStyle = 0
            pic_Edit.Visible = True
            Set pic_Edit.Container = pic_Left
            lbl_Left.Caption = "Text Editor"
        End If
        txt_Path.Text = f_GetPath
        lw_Right_ItemClick lw_Right.SelectedItem
    Case 5 ' MULTIMEDIA PREVIEW
        If Side = 1 Then
            lw_Left.Visible = False
            dir_Right.Visible = False
            pic_Search.Visible = False
            pic_Edit.Visible = False
            pic_Show.Visible = True
            pic_Left.BorderStyle = 0
            Set pic_Show.Container = pic_Left
            lbl_Left.Caption = "Multimedia Preview"
        End If
    End Select
    ' Activate the search if the search box is selected
    if_Size
    if_Size ' Twice, DO NOT REMOVE !
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_setmode": Resume Next
End Sub

Public Sub if_Size()
On Error Resume Next
    ' Why, now the boring part, this code is too large, so
    ' sometimes I need to call this twice, better that then
    ' changing the whole code...
    With pic_Right
        .Width = Me.ScaleWidth - .Left - 2
        .Height = Me.ScaleHeight - .Top - sb_StatusBar.Height - 2 - pic_DBar.Height
        lw_Right.Width = .Width - 5: lw_Right.Height = .Height - 5
        dir_Left.Width = .Width: dir_Right.Height = .Height
        lbl_Right.Left = .Left: lbl_Right.Width = .Width
    End With
    With pic_Left
        .Height = pic_Right.Height
        lbl_Left.Width = .Width
        lw_Left.Width = .Width - 5: lw_Left.Height = .Height - 5
        dir_Right.Width = .Width: dir_Right.Height = .Height
    End With
    With pic_SizeBar
        pic_Left.Width = .Left - pic_Left.Left
        pic_Right.Left = .Left + .Width
        pic_Right.Width = Me.ScaleWidth - pic_Right.Left - 2
    End With
    With pb_ProgressBar
        .Top = Me.ScaleHeight - .Height - 4
    End With
    With pic_Bar
        .Width = Me.ScaleWidth
        txt_Add.Width = .Width - txt_Add.Left - cmd_Fav.Width - 3
        cmd_Fav.Left = Me.ScaleWidth - cmd_Fav.Width - 1
        pic_Title.Width = .Width - 3
    End With
    With pic_Search
        If .Visible Then
            .Left = 0
            .Top = 0
            .Height = pic_Left.Height
            pic_Right.Left = .Left + .Width + 3
            pic_SizeBar.Left = .Left + .Width
        End If
    End With
    With pic_Edit
        If .Visible Then
            text_Fullscr False
            .Left = 0
            .Top = 0
            .Height = pic_Left.Height
            pic_Right.Left = .Left + .Width + 3
            pic_SizeBar.Left = .Left + .Width
            pic_Left.Width = .Width
            sb_Text.Top = .Height - sb_Text.Height
            sb_Text.Width = .Width
            txt_Edit.Height = .Height - txt_Edit.Top - sb_Text.Height
        End If
    End With
    With pic_Show
        If .Visible Then
            .Left = 0
            .Top = 0
            .Height = pic_Left.Height
            pic_Right.Left = .Left + .Width + 3
            pic_SizeBar.Left = .Left + .Width
        End If
    End With
End Sub




Private Sub sld_Memo_Change()
    ifMaxIconCashe = sld_Memo.Value
End Sub

Private Sub tb_Edit_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo e
    Dim I&
    With txt_Edit
        Select Case Button.Key
        Case "Save"
            If UCase(Left(if_LW.SelectedItem.Text, 3)) = "RTF" Then
                .SaveFile f_GetPath & "\" & if_LW.SelectedItem.Text, rtfRTF
            Else
                .SaveFile f_GetPath & "\" & if_LW.SelectedItem.Text, rtfText
            End If
        Case "Save As"
            cd_Edit.DialogTitle = "Save rich text file"
            cd_Edit.ShowSave
            If cd_Edit.Filename <> "" Then .SaveFile cd_Edit.Filename
        Case "New"
            .Text = ""
        Case "Find"
            text_Find
        Case "Font"
            cd_Edit.DialogTitle = "Select font for selected text"
            cd_Edit.fLags = cdlCFScreenFonts Or cdlCFEffects
            cd_Edit.FontName = ""
            cd_Edit.ShowFont
            If cd_Edit.FontName <> "" Then
                .SelBold = cd_Edit.FontBold
                .SelItalic = cd_Edit.FontItalic
                .SelFontName = cd_Edit.FontName
                .SelFontSize = cd_Edit.FontSize
                .SelStrikeThru = cd_Edit.FontStrikethru
                .SelUnderline = cd_Edit.FontUnderline
                .SelColor = cd_Edit.Color
            End If
        Case "Center"
            .SelAlignment = rtfCenter
        Case "Left"
            .SelAlignment = rtfLeft
        Case "Right"
            .SelAlignment = rtfRight
        Case "Word Wrap"
            If Button.Value = tbrPressed Then
                .RightMargin = 2500
            Else
                .RightMargin = 0 ' it will be same as .width
            End If
        Case "Key Commands"
            Dim msg$
            msg = "Extra commands: " & vbCrLf & "Undo " & vbTab & " CTRL+Z" & vbCrLf & "Cut " & vbTab & " CTR+X" & vbCrLf & "Copy " & vbTab & " CTRL+C" & vbCrLf & "Paste " & vbTab & " CTRL+V" & vbCrLf & "Delete " & vbTab & " DEL" & vbCrLf & "Find " & vbTab & " CTRL+F" & vbCrLf & "Select All " & vbTab & " CTRL+A"
            MsgBox msg, vbInformation, "Help"
        Case "Fullscreen"
            If Button.Value = tbrPressed Then
                text_Fullscr True
            Else
                if_Size
                text_Fullscr False
            End If
        End Select
    End With

Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "tb_edit_buttonclick": Resume Next
End Sub



Private Sub tb_Show_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo e
    Select Case Button.Key
    Case "Export"
        cd_Show.ShowOpen
        If Not cd_Show.Filename = "" Then
            pic_Save.Picture = img_Picture.Picture
            SavePicture pic_Save.Image, cd_Show.Filename
            pic_Save.Picture = LoadPicture("")
        End If
    Case "Full Screen Preview"
        With frm_Fullsrc
            .Show vbModal, Me
        End With
    Case "Image Properties"
        mnu_File_Props_Click
    Case "Next"
        image_Next
    Case "Previous"
        image_Previous
    Case "Zoom in"
        With img_Picture
            .Stretch = True
            .Width = .Width + 32
            .Height = .Height + 32
            ' Center the pic
            .Left = pic_Holder.Width / 2 - .Width / 2
            .Top = pic_Holder.Height / 2 - .Height / 2
        End With
    Case "Zoom out"
        With img_Picture
            If .Width > 64 Then
                .Stretch = True
                .Width = .Width - 32
                .Height = .Height - 32
                .Left = pic_Holder.Width / 2 - .Width / 2
                .Top = pic_Holder.Height / 2 - .Height / 2
            End If
        End With
    Case "Play"
        ' If u click it while this toolbar is visible, it will be the same
        lw_Right_ItemClick lw_Right.SelectedItem
    End Select
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "tb_show_buttonclick": Resume Next
End Sub









Private Sub tb_ToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Up One Level"
        mnu_File_Up_Click
    Case "Search"
        mnu_Show_FileSearch_Click
    Case "Text Edit"
        mnu_Show_Edit_Click
    Case "Multimedia"
        mnu_Show_FilePrev_Click
    Case "Directories"
        mnu_Show_Dirs_Click
    Case "Properties"
        mnu_File_Props_Click
    Case "System Info"
        mnu_Show_Sys_Click
    Case "Refresh"
        mnu_File_Refresh_Click
    End Select
End Sub

Private Sub txt_Add_Change()
    txt_Path.Text = txt_Add.Text
End Sub

Private Sub txt_Add_GotFocus()
    txt_Add.SelStart = 0
    txt_Add.SelLength = 999
End Sub

Private Sub txt_Add_KeyPress(KeyAscii As Integer)
On Error GoTo e
    If KeyAscii = 13 Then
        If Left(txt_Add.Text, 4) = "www." Then
            File_Open txt_Add.Text, "Open"
        Else
            if_Dir.Path = txt_Add.Text
            drv_Drive.Drive = if_Dir.Path
        End If
        KeyAscii = 0 ' Not to beep after
    End If
Exit Sub
e:
    MsgBox "The specified path could not be found !", vbExclamation, "Error"
    Resume Next
End Sub



Public Function if_GetIconCashe(Side As Integer) As Double
On Error GoTo e
    Dim Num As Double
    Select Case Side
    Case 1
        Num = img_ListView1.ListImages.Count
    Case 2
        Num = img_ListView2.ListImages.Count
    End Select
    if_GetIconCashe = (Num * 256) / 1024   ' This will get GDI memory used for icons (at least part of it)
Exit Function
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_geticoncashe": Resume Next
End Function

Public Function if_LW(Optional Side As Integer) As ListView
On Error GoTo e
    ' Set the selected ListView for if_LW
    If Side = 0 Then Side = ifSelected
    Select Case Side
    Case 1
        Set if_LW = lw_Left
    Case 2
        Set if_LW = lw_Right
    End Select
Exit Function
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_lw: " & Side: Resume Next
End Function

Public Function if_Dir(Optional Side As Integer) As DirListBox
On Error GoTo e
    ' Same as if_LW
    If Side = 0 Then Side = ifSelected
    Select Case Side
    Case 1
        Set if_Dir = dir_Left
    Case 2
        Set if_Dir = dir_Right
    End Select
Exit Function
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_dir: " & Side: Resume Next
End Function


Public Function if_ShowDirs() As String
    frm_Dirs.Show vbModal, Me
    if_ShowDirs = frm_Dirs.SelectedPath
End Function

Public Sub if_Refresh()
On Error GoTo e
    If ifSelected = 1 Then
        dir_Left.Refresh: Dir_Left_Change
    Else
        dir_Right.Refresh: dir_Right_Change
    End If
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_refresh": Resume Next
End Sub

Public Sub Zip_Open()
On Error GoTo e
    ' This will list all the files from a ZIP file onto LW
    Dim I&, zFile$, Itm As ListItem, Ext$, zIcon&, zName$, zPath$
    With if_LW
        zFile = .SelectedItem.Text ' Zip File
        if_ZipActivation True ' Activate the zip mode
        Open "C:\TICON.TMP" For Binary Access Write As #1 ' Creating the temp file... ;)
        Close #1
        .ListItems.Add(, , "<...>", , 1).SubItems(1) = "<DIR>" ' Add the up 1 level
        Zip_ReadArchive f_GetPath & "\" & zFile ' This will add all the zip files to a collection
        For I = 1 To Archive.Count ' Archive files count
            Ext = UCase(Right(Zip_GetEntry(I).Filename, 3)) ' Get extension
            ' Now this is a cheap trick
            Name "C:\TICON.TMP" As "C:\TICON." & Ext ' Set the temp file ext. to get an icon form it... :)
            zIcon = Icon_AddToImageList("C:\TICON." & Ext, Ext, img_ListView2)
            Name "C:\TICON." & Ext As "C:\TICON.TMP" ' Now return the old one
            zName = File_ParseName(Zip_GetEntry(I).Filename)
            zPath = File_ParsePath(Zip_GetEntry(I).Filename)
            If Not zName = "" Then ' It may be "", if you have used WinZip for creatig thi ZIP file
                Set Itm = .ListItems.Add(, , zName, , zIcon)
                With Itm
                    .SubItems(1) = Ext
                    .SubItems(2) = CLng(Zip_GetEntry(I).CompressedSize / 1024)
                    .SubItems(3) = Zip_GetEntry(I).CRC32
                    .SubItems(4) = Zip_GetEntry(I).FileDateTime
                    .SubItems(5) = CLng(Zip_GetEntry(I).UncompressedSize / 1024)
                    .SubItems(6) = zPath
                End With
            End If
        Next I
    End With
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "zip_open: " & zFile: Resume Next
End Sub

Public Sub if_ZipActivation(Active As Boolean)
On Error GoTo e
    ' Prepare the interface for ZIP archive
    mnu_File_Select.Enabled = Not Active
    mnu_File_TEdit.Enabled = Not Active
    mnu_File_Move.Enabled = Not Active
    mnu_File_Copy.Enabled = Not Active
    mnu_File_Rename.Enabled = Not Active
    mnu_File_CreateShortcut.Enabled = Not Active
    mnu_File_NewDir.Enabled = Not Active
    mnu_File_CreateDir.Enabled = Not Active
    mnu_File_Refresh.Enabled = Not Active
    mnu_File_Zip.Enabled = Not Active
    mnu_File_UnZip.Enabled = Active
    mnu_File_Props.Enabled = Not Active
    mnu_File_UzipSelected = Active
    With if_LW
        If Active Then
            .ListItems.Clear
            .ColumnHeaders.Clear
            With .ColumnHeaders
                .Add , , "File Name", 150
                .Add , , "Ext.", 50
                .Add , , "Size", 50
                .Add , , "CRC", 50
                .Add , , "Date", 120
                .Add , , "Unc. Size", 65
                .Add , , "Path", 500
            End With
        Else
            .ListItems.Clear
            .ColumnHeaders.Clear
            With .ColumnHeaders
                .Add , , "File Name", 150
                .Add , , "Ext.", 50
                .Add , , "Size", 80
                .Add , , "Modified", 120
                .Add , , "Attributes", 120
            End With
        End If
    End With
    ifZip = Active
Exit Sub
e:
    Err_Raise Err.Number, Err.Description, ModuleName, "if_zipactivation": Resume Next
End Sub

Private Sub txt_Edit_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    With sb_Text
        .Panels(1).Text = if_LW.SelectedItem.Text ' File Name
        .Panels(2).Text = "Line: " & Get_CurrentLine(txt_Edit) & " / " & Get_TotalLines(txt_Edit) & " Total "
    End With
    If Shift = 2 Then ' If CTRL holded
        If KeyCode = vbKeyF Then
            text_Find
        End If
    End If
End Sub


Private Sub txt_Edit_SelChange()
On Error Resume Next
    With sb_Text
        .Panels(1).Text = if_LW.SelectedItem.Text
        .Panels(2).Text = "Line: " & Get_CurrentLine(txt_Edit) & " / " & Get_TotalLines(txt_Edit) & " Total "
    End With
End Sub


Private Sub txt_Search_Change()
    If txt_Search.Text <> "" Then cmd_Search.Enabled = True Else cmd_Search.Enabled = False
End Sub

Private Sub txt_Size_Change()
    If Not IsNumeric(txt_Size.Text) Then txt_Size.Text = "0"
End Sub


Private Sub txt_Size2_Change()
    If Not IsNumeric(txt_Size2.Text) Then txt_Size2.Text = "0"
End Sub



Public Sub if_SearchActivation(Active As Boolean)
On Error Resume Next
    ' Prepare the interface for searching mode
    mnu_File_CreateShortcut.Enabled = Not Active
    mnu_File_NewDir.Enabled = Not Active
    mnu_File_CreateDir.Enabled = Not Active
    mnu_File_Refresh.Enabled = Not Active
    mnu_File_Zip.Enabled = Not Active
    mnu_File_UnZip.Enabled = Not Active
    txt_Add.Enabled = Not Active
    If Active Then
        txt_Search.SetFocus
    End If
End Sub

Public Sub text_Find()
    frm_TFind.Show vbModeless, Me
End Sub

Public Sub image_Next()
On Error Resume Next
    ' Just selects the next file
    With if_LW
        .ListItems(.SelectedItem.Index + 1).Selected = True
        .ListItems(.SelectedItem.Index - 1).Selected = False
        lw_Right_ItemClick .SelectedItem
    End With
End Sub

Public Sub image_Previous()
On Error Resume Next
    ' Previous file
    With if_LW
        .ListItems(.SelectedItem.Index - 1).Selected = True
        .ListItems(.SelectedItem.Index + 1).Selected = False
        lw_Right_ItemClick .SelectedItem
    End With
End Sub

Public Sub text_Fullscr(Full As Boolean)
    ' Show the fullscreen text editor
    If Full = True Then
        With pic_Left
            Me.WindowState = vbMaximized
            .Left = 0
            .Top = 0
            .ZOrder 0
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight
            pic_Edit.Width = .Width
            pic_Edit.Height = .Height
            sb_Text.Top = .Height - sb_Text.Height
            sb_Text.Width = .Width
            txt_Edit.Width = .Width
            txt_Edit.Height = .Height - .Top - sb_Text.Height * 2
        End With
        ifText = True
    Else
        With pic_Left
            If ifText = True Then
                .Left = 0
                .Top = 75
                .Width = 367
                txt_Edit.Width = 367
                pic_Edit.Width = .Width
                pic_Edit.Height = .Height
                tb_Edit.Buttons("Fullscreen").Value = tbrUnpressed
            End If
        End With
        ifText = False
    End If
    mnu_File.Enabled = Not Full
    mnu_Show.Enabled = Not Full
    mnu_Explorer.Enabled = Not Full
    mnu_Quick.Enabled = Not Full
End Sub
