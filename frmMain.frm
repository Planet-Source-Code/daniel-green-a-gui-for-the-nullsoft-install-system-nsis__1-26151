VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Slate Blue"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5520
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6932
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6156
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Text            =   "Slate Blue"
            TextSave        =   "Slate Blue"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstOptions 
      Height          =   5310
      ItemData        =   "frmMain.frx":1CCA
      Left            =   120
      List            =   "frmMain.frx":1CF5
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Contributors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   9
      Left            =   2160
      TabIndex        =   9
      Top             =   0
      Width           =   6255
      Begin VB.Label lblContributors 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1E1E
         Height          =   5055
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "License information &&"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin VB.Frame frameNSISLicenseInto 
         Caption         =   "Introduction to NSIS License && License File."
         Height          =   1335
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   6015
         Begin VB.CheckBox chkNSISLicense 
            Caption         =   "Insert a license into the script."
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Value           =   1  'Checked
            Width           =   3855
         End
         Begin VB.TextBox txtNSISLicense 
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Text            =   "license.txt"
            Top             =   960
            Width           =   5775
         End
         Begin VB.TextBox txtNSISLicenseIntro 
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Text            =   "This installer was created by Slate Blue and the user didn't change this text."
            Top             =   600
            Width           =   5775
         End
      End
      Begin VB.Frame frameNSISUserDir 
         Caption         =   "The text to prompt the user to enter a directory."
         Height          =   615
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   6015
         Begin VB.TextBox txtNSISUserDirPrompt 
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Text            =   "Please select your Winamp path below (you will be able to proceed when Winamp is detected):"
            Top             =   240
            Width           =   5775
         End
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Installation II - Output directory settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   6255
      Begin VB.Frame frameJeffsMom 
         Caption         =   "Path"
         Height          =   615
         Left            =   120
         TabIndex        =   64
         Top             =   1680
         Width           =   6015
         Begin VB.TextBox txtNSISInstallPathExtra 
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Text            =   "\Winamp"
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.CheckBox chkDetectWinamp 
         Caption         =   "Detect Winamp Directory using registry"
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   1440
         Width           =   3495
      End
      Begin VB.OptionButton optInstallDir 
         Caption         =   "Program Files (auto-detected)"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton optInstallDir 
         Caption         =   "Desktop"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optInstallDir 
         Caption         =   "Windows Directory (auto-detected)"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   60
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton optInstallDir 
         Caption         =   "Windows System Directory (auto-detected)"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   59
         Top             =   960
         Width           =   3495
      End
      Begin VB.OptionButton optInstallDir 
         Caption         =   "System Temp. Directory (auto-detected)"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   1200
         Width           =   3255
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "About Slate Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   8
      Left            =   2160
      TabIndex        =   8
      Top             =   0
      Width           =   6255
      Begin VB.Frame Frame2 
         Caption         =   "Credits"
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   6015
         Begin VB.Label lblLiquid 
            Caption         =   "Creator of AVS (source of the GUI of this GUI) - Nullsoft"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   480
            Width           =   5775
         End
         Begin VB.Label lblLiquid 
            Caption         =   "Learner of NSIS scripting language, coder of GUI - dan green"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   " Version:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   " MorphedMedia"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "NSIS scripting for the masses."
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Slate Blue:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label lblVer 
         BackStyle       =   0  'Transparent
         Caption         =   "Slate Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblIntro 
         Caption         =   "Copyright (C) 2001"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Scripting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Top             =   0
      Width           =   6255
      Begin VB.Frame frameNSISTitle 
         Caption         =   "Installation Title (ex. TinyVis, AVS)"
         Height          =   615
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   6015
         Begin VB.TextBox txtNSISName 
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Text            =   "Slate Blue Default"
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameNSISOutputName 
         Caption         =   "NSIS Output Filename (ex. TinyVis.exe)"
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   6015
         Begin VB.TextBox txtNSISOutput 
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Text            =   "SlateBlueDef.exe"
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameNSISMisc 
         Caption         =   "Miscellaneous"
         Height          =   615
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   6015
         Begin VB.CheckBox chkNSISCRC 
            Caption         =   "Perform CRC check on installer start."
            Height          =   210
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   3855
         End
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "To Do List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   12
      Left            =   2160
      TabIndex        =   50
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtScheduledWork 
         BackColor       =   &H80000004&
         Height          =   5055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "What's New?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   11
      Left            =   2160
      TabIndex        =   47
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtWhatsNew 
         BackColor       =   &H80000004&
         Height          =   5055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Installation I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   4
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   6255
      Begin VB.Frame frameInstallOptions 
         Caption         =   "Section Name"
         Height          =   615
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   6015
         Begin VB.TextBox txtNSISSectionName 
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Text            =   "ThisNameIsIgnoredSoWhyBother?"
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameOutPath 
         Caption         =   "Change the OutPath (added for functionality)"
         Height          =   615
         Left            =   120
         TabIndex        =   54
         Top             =   840
         Width           =   6015
         Begin VB.TextBox txtNSISOutPath 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Text            =   "$INSTDIR"
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameShortcuts 
         Caption         =   "Shortcuts"
         Height          =   3855
         Left            =   120
         TabIndex        =   89
         Top             =   1440
         Width           =   6015
         Begin VB.CheckBox chkShortcuts 
            Caption         =   "Create shortcuts on installation."
            Height          =   210
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Value           =   1  'Checked
            Width           =   5775
         End
         Begin VB.ListBox lstNSISShortcuts 
            Appearance      =   0  'Flat
            Height          =   1470
            ItemData        =   "frmMain.frx":1F3A
            Left            =   120
            List            =   "frmMain.frx":1F3C
            Style           =   1  'Checkbox
            TabIndex        =   90
            Top             =   480
            Width           =   5775
         End
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Add files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   0
      Width           =   6255
      Begin VB.DirListBox dirAdd 
         Height          =   330
         Left            =   120
         TabIndex        =   46
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.FileListBox fileAdd 
         Height          =   300
         Left            =   1080
         TabIndex        =   45
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdAddDir 
         Caption         =   "Add Folder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   41
         Top             =   240
         Width           =   2055
      End
      Begin VB.ListBox Container 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   600
         TabIndex        =   25
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog dialog 
         Left            =   120
         Top             =   4200
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Add file..."
      End
      Begin VB.ListBox lstInputFiles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4350
         ItemData        =   "frmMain.frx":1F3E
         Left            =   120
         List            =   "frmMain.frx":1F40
         TabIndex        =   24
         Top             =   600
         Width           =   6015
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "Remove"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdRemoveAll 
         Caption         =   "Remove all"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   5040
         Width           =   1335
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Compiler Settings"
      Height          =   5415
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CheckBox chkCheckSyntax 
         Caption         =   "Check syntax before compiling. (soon!)"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   3855
      End
      Begin VB.Frame frameNSISPath 
         Caption         =   "NSIS Path (auto-detected)"
         Height          =   615
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   6015
         Begin VB.TextBox txtNSISPath 
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameCompilerCmdLine 
         Caption         =   "Command Line options"
         Height          =   615
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   6015
         Begin VB.CheckBox chkSlashPause 
            Caption         =   "/Pause - Compiler pauses after script compilation"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Value           =   1  'Checked
            Width           =   3855
         End
      End
      Begin VB.CheckBox chkAddComments 
         Caption         =   "Comment Slate Blue-compiled scripts."
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkGenerateBeforeCompile 
         Caption         =   "Generate the script before running the compiler."
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Value           =   1  'Checked
         Width           =   3855
      End
   End
   Begin VB.Frame frameOptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox picLogo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   1080
         Picture         =   "frmMain.frx":1F42
         ScaleHeight     =   2235
         ScaleWidth      =   4035
         TabIndex        =   20
         Top             =   1680
         Width           =   4095
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   3
      Left            =   2160
      TabIndex        =   67
      Top             =   0
      Width           =   6255
      Begin VB.Frame frameInstallInfoCOlors 
         Caption         =   "Install Info Colors"
         Height          =   1215
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   6015
         Begin VB.Frame frameInstallInfoCOlors 
            Caption         =   "Colors"
            Height          =   615
            Index           =   1
            Left            =   120
            TabIndex        =   79
            Top             =   480
            Width           =   5775
            Begin VB.PictureBox picInstallColor 
               Height          =   255
               Index           =   1
               Left            =   5280
               ScaleHeight     =   195
               ScaleWidth      =   315
               TabIndex        =   85
               Top             =   240
               Width           =   375
            End
            Begin VB.PictureBox picInstallColor 
               Height          =   255
               Index           =   0
               Left            =   2280
               ScaleHeight     =   195
               ScaleWidth      =   315
               TabIndex        =   84
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtInstallColor 
               Height          =   255
               Index           =   0
               Left            =   1560
               MaxLength       =   6
               TabIndex        =   81
               Text            =   "FF0800"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtInstallColor 
               Height          =   255
               Index           =   1
               Left            =   4560
               MaxLength       =   6
               TabIndex        =   80
               Text            =   "000030"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblInstallColor 
               Caption         =   "Foreground Color: #"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   83
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label lblInstallColor 
               Caption         =   "Background Color: #"
               Height          =   255
               Index           =   1
               Left            =   3000
               TabIndex        =   82
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.CheckBox chkColorWindows 
            Caption         =   "Use windows default colors."
            Height          =   210
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame frameGradient 
         Caption         =   "Gradient Background Window"
         Height          =   1215
         Left            =   120
         TabIndex        =   70
         Top             =   1440
         Width           =   6015
         Begin VB.CheckBox chkGradientBG 
            Caption         =   "Use a gradient background window."
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   3855
         End
         Begin VB.Frame frameColors 
            Caption         =   "Colors"
            Height          =   615
            Left            =   120
            TabIndex        =   71
            Top             =   480
            Width           =   5775
            Begin VB.PictureBox picBGColor 
               Height          =   255
               Index           =   2
               Left            =   5280
               ScaleHeight     =   195
               ScaleWidth      =   315
               TabIndex        =   88
               Top             =   240
               Width           =   375
            End
            Begin VB.PictureBox picBGColor 
               Height          =   255
               Index           =   1
               Left            =   3360
               ScaleHeight     =   195
               ScaleWidth      =   315
               TabIndex        =   87
               Top             =   240
               Width           =   375
            End
            Begin VB.PictureBox picBGColor 
               Height          =   255
               Index           =   0
               Left            =   1320
               ScaleHeight     =   195
               ScaleWidth      =   315
               TabIndex        =   86
               Top             =   240
               Width           =   375
            End
            Begin VB.TextBox txtNSISBGGradientColor 
               Height          =   255
               Index           =   0
               Left            =   600
               MaxLength       =   6
               TabIndex        =   74
               Text            =   "000000"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtNSISBGGradientColor 
               Height          =   255
               Index           =   1
               Left            =   2640
               MaxLength       =   6
               TabIndex        =   73
               Text            =   "800000"
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtNSISBGGradientColor 
               Height          =   255
               Index           =   2
               Left            =   4560
               MaxLength       =   6
               TabIndex        =   72
               Text            =   "FFFFFF"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblTopColor 
               Caption         =   "Top: #"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   77
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lblTopColor 
               BackStyle       =   0  'Transparent
               Caption         =   "Bottom: #"
               Height          =   255
               Index           =   1
               Left            =   1920
               TabIndex        =   76
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblTopColor 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Text:#"
               Height          =   255
               Index           =   2
               Left            =   3960
               TabIndex        =   75
               Top             =   240
               Width           =   495
            End
         End
      End
   End
   Begin VB.Frame frameOptions 
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Index           =   10
      Left            =   2160
      TabIndex        =   10
      Top             =   0
      Width           =   6255
      Begin VB.Label lblGreetz 
         Caption         =   $"frmMain.frx":4211
         Height          =   5055
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoadPreset 
         Caption         =   "&Load Script"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Script"
      End
      Begin VB.Menu mnuFileSpacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewEditor 
         Caption         =   "&Show Script Editor"
      End
   End
   Begin VB.Menu mnuComple 
      Caption         =   "&Compile"
      Begin VB.Menu mnuCompileGenerate 
         Caption         =   "&Generate Script"
      End
      Begin VB.Menu mnuCompileGo 
         Caption         =   "&Go"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "&Mode"
      Enabled         =   0   'False
      Begin VB.Menu mnuModeGUI 
         Caption         =   "&GUI Script Editing"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuModeManual 
         Caption         =   "&Manual Script Editing"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpSpacer0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpToDoList 
         Caption         =   "&To Do List"
      End
      Begin VB.Menu mnuHelpWhatsNew 
         Caption         =   "&What's New"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
'Slate Blue for NSIS
'Written by Dan Green
'Greetz to Nullsoft
'Copyright? 2001 Dan Green
'visit morphedmedia.com/slateblue for more information
'note: i've tried to fully comment the source, but it's
'a work in progress, so it might not be 100%. :)
'===========================================================
'code: i've used the error trapping technique of:
'On Error Resume Next
'in about every sub or function, as it seems to work the
'best in providing a seemingly errorless program(to the user
'at least).
Option Explicit
'used in exiting the program
Dim boolEnd As Boolean
'contains the $path variables for NSIS
Dim strNSIS_INSTALL_PATHS(4) As String
'contains the NSIS path,the path to the current script
Dim strNSIS_PATH As String, strNSIS_SCRIPT_PATH As String
'used in parsing
Dim strBuffer As String, strBuffer2 As String, strBuffer3 As String, intStart As Integer, intEnd As Integer, intTitle As Integer
'contains the lame stuff i haven't had time to properly tag
Dim retval, a, b, c, i, k, Filename
'used for executing NSIS
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub chkColorWindows_Click()
If chkColorWindows.Value = 1 Then
txtInstallColor(0).Enabled = False
txtInstallColor(1).Enabled = False
lblInstallColor(0).Enabled = False
lblInstallColor(1).Enabled = False
Else
txtInstallColor(0).Enabled = True
txtInstallColor(1).Enabled = True
lblInstallColor(0).Enabled = True
lblInstallColor(1).Enabled = True
End If
End Sub

Private Sub chkGradientBG_Click()
If chkGradientBG.Value = 0 Then
txtNSISBGGradientColor(0).Enabled = False
txtNSISBGGradientColor(1).Enabled = False
txtNSISBGGradientColor(2).Enabled = False
lblTopColor(0).Enabled = False
lblTopColor(1).Enabled = False
lblTopColor(2).Enabled = False
Else
txtNSISBGGradientColor(0).Enabled = True
txtNSISBGGradientColor(1).Enabled = True
txtNSISBGGradientColor(2).Enabled = True
lblTopColor(0).Enabled = True
lblTopColor(1).Enabled = True
lblTopColor(2).Enabled = True
End If
End Sub

Private Sub cmdAdd_Click()
'this sub add a file or files to the script. these files are
'the ones that get installed by the compiled script.
On Error Resume Next
'set the filer
dialog.Filter = "All Files (*.*) | *.*"
' Show the Open Dialog
dialog.ShowOpen
' Check to see if the user selected a file
If dialog.Filename = "" Then Exit Sub
' See if the file was already added
For i = 0 To lstInputFiles.ListCount - 1
    If lstInputFiles.List(i) = dialog.Filename Then Exit Sub
Next i
' Now we need to make sure that the file isn't empty
' If an error occurs, the file doesn't exist
On Error GoTo NoFile
' Check to see if the file has a a size of 0
If FileLen(dialog.Filename) <= 0 Then
    ' Display a Yes-No Box asking the user if he would
    ' still like to add the file even though it has no
    ' content
    retval = MsgBox("The file " & dialog.Filename & " has a zero Byte length (Its Empty)!" & _
                    vbNewLine & "Are you Sure you want to add it?", vbYesNo, "Error")
    ' User clicked No
    If retval = vbNo Then
        Exit Sub
    End If
End If
' Now add the file to the list boxes
lstInputFiles.AddItem dialog.Filename
lstNSISShortcuts.AddItem dialog.Filename
Container.AddItem dialog.FileTitle
' Enable the Next button
cmdRemove.Enabled = True
cmdRemoveAll.Enabled = True

NoFile:
End Sub

Private Sub cmdRemove_Click()
'this sub removes all selected files from the script.
On Error Resume Next
' Scan through each item in the listbox to see if its selected
For i = 0 To lstInputFiles.ListCount - 1
    If lstInputFiles.Selected(i) Then
        ' Remove the selected Item
        lstInputFiles.RemoveItem i
        Container.RemoveItem i
        lstNSISShortcuts.RemoveItem i
        ' Now check to see if there are any more items in the
        ' Listboxes
        If lstInputFiles.ListCount = 0 Then
            ' If there aren't, disable the remove buttons
            cmdRemove.Enabled = False
            cmdRemoveAll.Enabled = False
        Else
            ' if there are, Enable the remove buttons
            
        End If
        Exit Sub
    End If
Next i
End Sub

Private Sub cmdRemoveAll_Click()
'this sub removes all files from the script
On Error Resume Next
'part of the old slate blue that lies dormant.
Container.Clear
'list of files in the script
lstInputFiles.Clear
'list of shortcuts
lstNSISShortcuts.Clear
'disable the remove buttons (you can't remove if there are no
'files to remove!)
cmdRemove.Enabled = False
cmdRemoveAll.Enabled = False
End Sub

Private Sub cmdAddDir_Click()
'this sub shows the add directory form.
On Error Resume Next
frmAddDir.Show
End Sub

Private Sub Form_Load()
'MsgBox FileLen("c:\program files\microsoft visual studio\vb98\nsis gui\slateblue.exe")
'the all-important form_load()
On Error Resume Next
'set end program to no
boolEnd = False
'set color options
picInstallColor(0).BackColor = Val("&H" & txtInstallColor(0).text)
picInstallColor(1).BackColor = Val("&H" & txtInstallColor(1).text)
picBGColor(0).BackColor = Val("&H" & txtNSISBGGradientColor(0).text)
picBGColor(1).BackColor = Val("&H" & txtNSISBGGradientColor(1).text)
picBGColor(2).BackColor = Val("&H" & txtNSISBGGradientColor(2).text)
'set all option frames font to Arial
For i = 0 To frameOptions.Count - 1
    frameOptions(i).FontName = "arial"
Next i
'add a file to the add file section for default usage
lstInputFiles.AddItem App.path & "\slateblue.exe"
'set-up NSIS install path variables
strNSIS_INSTALL_PATHS(0) = "$PROGRAMFILES"
strNSIS_INSTALL_PATHS(1) = "$DESKTOP"
strNSIS_INSTALL_PATHS(2) = "$WINDIR"
strNSIS_INSTALL_PATHS(3) = "$SYSDIR"
strNSIS_INSTALL_PATHS(4) = "$TEMP"
'should load from registry, but it's all messed up
'strNSIS_PATH = GetSettingString(HKEY_CLASSES_ROOT, "NSISFile\DefaultIcon", (Default))
'that needs a bugfix
strNSIS_PATH = "C:\Program Files\NSIS\makensis.exe"
txtNSISPath = strNSIS_PATH
'sets the version label = the app versions
lblVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
'make sure the status bar is ready
Status.Panels(1).text = "Ready"
Status.Panels(2).text = "No File Loaded..."
Status.Panels(3).text = "Slate Blue"
'loads the previous script
'===LoadData
'load the scheduled work data
txtScheduledWork.text = ReadFile(App.path & "\to do.txt")
'load the What's New? data
txtWhatsNew.text = ReadFile(App.path & "\whats new.txt")
'make sure the first frame is showing
lstOptions.Selected(0) = True
lstOptions_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this sub sets the mouseover color of our www url to blue
On Error Resume Next
Label3.ForeColor = vbBlue
End Sub

Private Sub Form_Unload(Cancel As Integer)
'unloads the program
On Error Resume Next
'saves the script to a temp file
'===SaveData
'make sure all loops stop
boolEnd = True
'make sure the program quits, not just hides
Cancel = 0
'unload the form
Unload frmScript
'unload the form
Unload frmAddDir
'unload the form
Unload Me
End Sub

Private Sub frameOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the mouseover color of our www url to blue
On Error Resume Next
Label3.ForeColor = vbBlue
End Sub

Private Sub Label3_Click()
'opens the link our www url provides
On Error Resume Next
ShellExecute Me.hwnd, "Open", "http://www.morphedmedia.com", vbNullString, vbNullString, 0
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the mouseover color of our www url to red
On Error Resume Next
Label3.ForeColor = vbRed
End Sub

Private Sub lblIntro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the mouseover color of our www url to blue
On Error Resume Next
Label3.ForeColor = vbBlue
End Sub

Private Sub lblVer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the mouseover color of our www url to blue
On Error Resume Next
Label3.ForeColor = vbBlue
End Sub

Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'sets the mouseover color of our www url to blue
On Error Resume Next
Label3.ForeColor = vbBlue
End Sub




Private Sub lstOptions_Click()
'this sub is our navigational code
On Error Resume Next
'make all frames invisible
For i = 0 To frameOptions.Count - 1
    frameOptions(i).Visible = False
Next i
'get the frame we need
    i = lstOptions.ListIndex
'show the user-selected frame
Status.Panels(1).text = "Now Viewing: " & frameOptions(i).Caption
'on/off buttons
If i = 2 Then
    If lstInputFiles.ListCount >= 1 Then
        cmdRemove.Enabled = True
        cmdRemoveAll.Enabled = True
    Else
        cmdRemove.Enabled = False
        cmdRemoveAll.Enabled = False
    End If
End If
'show frame
frameOptions(i).Visible = True
End Sub

Private Sub mnuCompileGenerate_Click()
'this sub generates the script based on user input
'the "If chkAddComments.Value = 1" statements
'make our compiler check to see if it should add
'comments to the script, making the script easier
'to understand.
'the use of double-quotes is necessary in NSIS to
'make the script work.  to do this i simply add "chr(34)"
'where i need to double-quote (w/o the quotes).
On Error Resume Next
'clear the current script
frmScript.txtNSISScript.text = ""
'start the status bar
Status.Panels(1).text = "Generating Script..."
'insert Slate Blue Header
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & "; This script created by Slate Blue, a product of morphedmedia.com." & vbNewLine
End If
'insert install info colors
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & "; Install Info Colors" & vbNewLine
End If
If chkColorWindows.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "InstallColors /windows"
Else
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "InstallColors " & txtInstallColor(0).text & " " & txtInstallColor(1).text
End If
'insert the bg gradient color data
If chkGradientBG.Value = 1 Then
    If chkAddComments.Value = 1 Then
        frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; Setup window options & colors" & vbNewLine
    End If
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "BGGradient " & txtNSISBGGradientColor(0) & " " & txtNSISBGGradientColor(1) & " " & txtNSISBGGradientColor(2)
End If
'insert the title
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; Title of this installation"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "Name " & Chr(34) & txtNSISName & Chr(34)
'insert crc-check data
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & ";Perform a CRC check on the installer .exe at runtime"
End If
If chkNSISCRC.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "CRCCheck on"
Else
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "CRCCheck off"
End If
'insert the output filename
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & ";Output Filename"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "OutFile " & Chr(34) & txtNSISOutput & Chr(34)
'insert license introduction
If chkNSISLicense.Value = 1 Then
    If chkAddComments.Value = 1 Then
        frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; License Page Introduction"
    End If
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "LicenseText " & Chr(34) & txtNSISLicenseIntro & Chr(34)
'insert license data
    If chkAddComments.Value = 1 Then
        frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; License Data"
    End If
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "LicenseData " & txtNSISLicense
End If
'insert installation directory data
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; The default installation directory"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "InstallDir " & GetDirSelection & txtNSISInstallPathExtra.text
'insert detect winamp directory data
If chkDetectWinamp.Value = 1 Then
    If chkAddComments.Value = 1 Then
        frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; Detect winamp directory if available"
    End If
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "InstallDirRegKey HKLM" & " \ " & Chr(34) & "Software\Microsoft\Windows\CurrentVersion\Uninstall\Winamp" & Chr(34) & " \ " & Chr(34) & "UninstallString" & Chr(34)
End If
'insert prompt user for install directory data
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; The text to prompt the user to enter a directory"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "DirText " & Chr(34) & txtNSISUserDirPrompt.text & Chr(34)
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "DirShow hide"
'insert install data & create new section
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; The stuff to install"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "Section " & Chr(34) & txtNSISSectionName.text & Chr(34)
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; Set output path = already-chosen installation dir"
End If
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "SetOutPath $INSTDIR"
If chkAddComments.Value = 1 Then
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; Put the file(s) there"
End If
'add files to script
For i = 0 To lstInputFiles.ListCount - 1
    If chkAddComments.Value = 1 Then
        frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; File Number: " & (i + 1)
    End If
    frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "File " & Chr(34) & lstInputFiles.List(i) & Chr(34)
Next i
'end the installation files section
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "SectionEnd"
'signal end of file
frmScript.txtNSISScript.text = frmScript.txtNSISScript.text & vbNewLine & "; eof"
'stop the status bar
Status.Panels(1).text = "Script Generation Complete!"
End Sub

Private Sub mnuCompileGo_Click()
'this sub runs nsis on the script
'checks to see if the script should (re)generated before
'actually compiling
If chkGenerateBeforeCompile.Value = 1 Then
mnuCompileGenerate_Click
End If
'save the script before compiling
mnuFileSave_Click
'navigate to the slate blue logo screen
lstOptions.Selected(6) = True
lstOptions_Click
'this took a while, but i got it: this runs NSIS with
'our script and the user defined switches
retval = ShellExecute(Me.hwnd, "Open", strNSIS_PATH, "/CD " & GetCmdLine & " " & Chr(34) & strNSIS_SCRIPT_PATH & Chr(34), vbNullString, 1)
End Sub

Private Sub mnuFileExit_Click()
'this sub exits the program
On Error Resume Next
boolEnd = True
Form_Unload 0
End Sub

Private Sub mnuFileLoadPreset_Click()
'this sub loads a script
On Error Resume Next
'make sure commondialog can only see scripts
dialog.Filter = "NSIS Scripts (*.nsi) | *.nsi"
'commondialog will open in the NSIS path
dialog.InitDir = strNSIS_PATH
'show the commondialog open box
dialog.ShowOpen
'if there is no file chosen, or cancel was selected, exit sub
If dialog.Filename = "" Then Exit Sub
'if the file is empty exit sub
If Not FileLen(dialog.Filename) > 10 Then Exit Sub
'set the script_path variable
strNSIS_SCRIPT_PATH = dialog.Filename
'read the script so it can be parsed
frmScript.txtParse.text = ReadFile(dialog.Filename)
'parse the script
Parse frmScript.txtParse
'read the script into the script box
frmScript.txtNSISScript.text = ReadFile(dialog.Filename)
'set the status bar
Status.Panels(2).text = GetDir1(dialog.Filename)
End Sub

Private Sub mnuFileSave_Click()
'this sub saves the script
On Error Resume Next
'used for our file access
Dim intFileNum As Integer
'make sure commondialog can only see scripts
dialog.Filter = "NSIS Scripts (*.nsi) | *.nsi"
'show the commondialog save box
dialog.ShowSave
'if there is no file chosen, or cancel was selected, exit sub
If dialog.Filename = "" Then Exit Sub
'set the script_path variable
strNSIS_SCRIPT_PATH = dialog.Filename
'open the file, Print script to the file, & close the file
intFileNum = FreeFile
Open strNSIS_SCRIPT_PATH For Output As #intFileNum
Print #intFileNum, frmScript.txtNSISScript.text
Close #intFileNum
End Sub

Public Function GetDirSelection()
'this function returns which path variable is selected
On Error Resume Next
'read each variable until the selected is found
For i = 0 To optInstallDir.Count - 1
If optInstallDir(i).Value = True Then
GoTo 300
End If
Next i
'return the selected variable
300: GetDirSelection = strNSIS_INSTALL_PATHS(i)
End Function

Public Function GetCmdLine()
'this functions returns the command line switchs/variables
'only one so far, /PAUSE is the one we need most
If chkSlashPause.Value = 1 Then
GetCmdLine = "/PAUSE"
Else
GetCmdLine = ""
End If
Exit Function
End Function

Public Sub DoDirs(DirPath As String, DirFilters As String)
'this sub provides the ability to recurse subdirectories/files
'this remains uncommented to for now
On Error Resume Next
    fileAdd.Pattern = DirFilters
    dirAdd.path = DirPath
   DoFiles DirPath
        If dirAdd.ListCount = 0 Then Exit Sub
     For k = 0 To dirAdd.ListCount - 1
            dirAdd.path = DirPath
         DoDirs dirAdd.List(k), DirFilters
         'DoEvents allows the program to function while it's working
                DoEvents
            Next k
            dirAdd.path = DirPath
        End Sub

Private Sub DoFiles(DirPath As String)
'this sub is part of the recursive subdirectory sub
On Error Resume Next
    fileAdd.path = DirPath
    If fileAdd.ListCount = 0 Then Exit Sub
   For k = 0 To fileAdd.ListCount - 1
        Filename = fileAdd.path & String(1 - Abs(CInt(Right(fileAdd.path, 1) = "\")), "\") & fileAdd.List(k)
       lstInputFiles.AddItem fileAdd.path & "\" & Filename
    Next k
End Sub

Private Sub SaveData()
'this sub saves the script to a temp file.  some stupid bug,
'where the quotes multiply at runtime, forces me to disable
'this sub until i can resolve the error.
On Error Resume Next
'loads a string
Dim tmpFilename As String, intFileNum As Integer
'set the string = the temp script filename
tmpFilename = App.path & "\slateblue.dat"
'open the file, Print the header & script, & close the file
intFileNum = FreeFile
Open tmpFilename For Output As intFileNum
Print #intFileNum, frmScript.txtNSISScript
Close #intFileNum
End Sub

Private Sub LoadData()
'this sub loads the previous script.  some stupid bug,
'where the quotes multiply at runtime, forces me to disable
'this sub until i can resolve the error.
On Error Resume Next
If FileLen(App.path & "\slateblue.dat") = 0 Then Exit Sub
frmScript.txtNSISScript.text = ReadFile(App.path & "\slateblue.dat")
End Sub

Private Sub mnuHelpAbout_Click()
lstOptions.Selected(7) = True
lstOptions_Click
End Sub

Private Sub mnuHelpToDoList_Click()
lstOptions.Selected(11) = True
lstOptions_Click
End Sub

Private Sub mnuHelpWhatsNew_Click()
lstOptions.Selected(10) = True
lstOptions_Click
End Sub

Private Sub Parse(script As TextBox)
On Error Resume Next
'get the installation name
strBuffer3 = script.text
ClearVars
intTitle = InStr(script.text, "Name ")
intStart = InStr(intTitle, script.text, Chr(34), vbTextCompare)
If (intStart - intTitle) <= 5 Then
intEnd = InStr(intStart + 1, script.text, Chr(34), vbTextCompare)
strBuffer = Mid(script.text, (intStart + 1), ((intEnd - intStart) - 1))
Else
strBuffer = ""
End If
txtNSISName.text = strBuffer
'get the installation exe
ClearVars
intTitle = InStr(script.text, "OutFile ")
intStart = InStr(intTitle, script.text, Chr(34), vbTextCompare)
If (intStart - intTitle) <= 8 Then
intEnd = InStr(intStart + 1, script.text, Chr(34), vbTextCompare)
strBuffer = Mid(script.text, (intStart + 1), ((intEnd - intStart) - 1))
Else
strBuffer = ""
End If
strBuffer3 = Mid(strBuffer3, (intTitle + intEnd), (Len(strBuffer3) - (intTitle + intEnd)))
txtNSISOutput.text = strBuffer
'get the license text
ClearVars
intTitle = InStr(script.text, "LicenseText ")
intStart = InStr(intTitle, script.text, Chr(34), vbTextCompare)
If (intStart - intTitle) <= 5 Then
intEnd = InStr(intStart + 1, script.text, Chr(34), vbTextCompare)
Else
strBuffer = ""
End If
strBuffer = Mid(script.text, (intStart + 1), ((intEnd - intStart) - 1))
txtNSISLicenseIntro.text = strBuffer
'get the license data = complex, reads the line, not quotes
ClearVars
intTitle = InStr(script.text, "LicenseData")
intStart = (intTitle + Len("LicenseData "))
If (intStart - intTitle) <= 5 Then
intEnd = InStr((intStart), script.text, vbNewLine, vbTextCompare)
strBuffer = Mid(script.text, (intStart), ((intEnd - intStart)))
Else
strBuffer = ""
End If
txtNSISLicense.text = strBuffer
'get the prompt user to enter a dir text
ClearVars
intTitle = InStr(script.text, "DirText")
intStart = (intTitle + Len("DirText "))
If (intStart - intTitle) <= 5 Then
intEnd = InStr((intStart), script.text, vbNewLine, vbTextCompare)
strBuffer = Mid(script.text, (intStart + 1), ((intEnd - intStart) - 2))
Else
strBuffer = ""
End If
txtNSISUserDirPrompt.text = strBuffer
'get the install dir
ClearVars
intTitle = InStr(script.text, "InstallDir")
intStart = (intTitle + Len("InstallDir "))
intEnd = InStr((intStart), script.text, vbNewLine, vbTextCompare)
strBuffer = Mid(script.text, (intStart), ((intEnd - intStart)))
intStart = InStr(strBuffer, "\")
strBuffer2 = Mid(strBuffer, 2, (intStart - 2))
strBuffer = Right(strBuffer, (Len(strBuffer) - Len(strBuffer2) - 2))
GetInstallVariable (TrimQuotes(strBuffer2))
If Left(TrimQuotes(strBuffer), 1) <> "\" Then
strBuffer = "\" & strBuffer
End If
txtNSISInstallPathExtra.text = TrimQuotes(strBuffer)
'get the install files section name
ClearVars
intTitle = InStr(script.text, "Section")
intStart = InStr(intTitle, script.text, Chr(34), vbTextCompare)
If (intStart - intTitle) <= 8 Then
intEnd = InStr(intStart + 1, script.text, Chr(34), vbTextCompare)
strBuffer = Mid(script.text, (intStart + 1), ((intEnd - intStart) - 1))
Else
strBuffer = ""
End If
txtNSISSectionName.text = strBuffer
'get the second install dir
ClearVars
intTitle = InStr(script.text, "SetOutPath")
intStart = (intTitle + Len("SetOutPath"))
If (intStart - intTitle) <= 10 Then
intEnd = InStr((intStart), script.text, vbNewLine, vbTextCompare)
strBuffer = Trim(Mid(script.text, (intStart), ((intEnd - intStart))))
Else
strBuffer = ""
End If
txtNSISOutPath.text = strBuffer
'get the crc-check data
ClearVars
intTitle = InStr(script.text, "CRCCheck")
intStart = (intTitle + Len("CRCCheck "))
If (intStart - intTitle) <= 9 Then
intEnd = InStr((intStart), script.text, vbNewLine, vbTextCompare)
strBuffer = Mid(script.text, (intStart), ((intEnd - intStart)))
Else
strBuffer = ""
End If
If LCase(strBuffer) = "on" Then
chkNSISCRC.Value = 1
Else
chkNSISCRC.Value = 0
End If
'shit: reading multiple files from script into listbox
ClearVars
lstInputFiles.Clear
intTitle = InStr(strBuffer3, "File " & Chr(34))
intStart = InStr(intTitle, strBuffer3, Chr(34), vbTextCompare)
Do While Not boolEnd
Do While intTitle <> 0
If (intStart - intTitle) <= 73 Then
intEnd = InStr(intStart + 1, strBuffer3, Chr(34), vbTextCompare)
strBuffer = Mid(strBuffer3, (intStart + 1), ((intEnd - intStart) - 1))
Else
strBuffer = ""
End If
If CheckFilename(strBuffer) = True Then
lstInputFiles.AddItem strBuffer
End If
strBuffer3 = Mid(strBuffer3, (intEnd - intTitle), (Len(strBuffer3) - (intEnd - intTitle)))
intTitle = InStr(strBuffer3, "File " & Chr(34))
intStart = InStr(intTitle, strBuffer3, Chr(34), vbTextCompare)
DoEvents
Loop
boolEnd = True
Loop
End Sub

Private Sub GetInstallVariable(strText As String)
If strText = strNSIS_INSTALL_PATHS(0) Then
optInstallDir(0).Value = True
Exit Sub
End If
If strText = strNSIS_INSTALL_PATHS(1) Then
optInstallDir(1).Value = True
Exit Sub
End If
If strText = strNSIS_INSTALL_PATHS(2) Then
optInstallDir(2).Value = True
Exit Sub
End If
If strText = strNSIS_INSTALL_PATHS(3) Then
optInstallDir(3).Value = True
Exit Sub
End If
If strText = strNSIS_INSTALL_PATHS(4) Then
optInstallDir(4).Value = True
Exit Sub
End If
End Sub

Private Sub ClearVars()
intStart = 0
intEnd = 0
intTitle = 0
strBuffer = ""
strBuffer2 = ""
boolEnd = False
End Sub

Private Function CheckFilename(Filename As String) As Boolean
For i = 0 To lstInputFiles.ListCount - 1
    If lstInputFiles.List(i) = Filename Then
    CheckFilename = False
    Exit Function
    End If
Next i
CheckFilename = True
End Function

Private Sub mnuViewEditor_Click()
If mnuViewEditor.Checked = True Then
frmScript.Hide
mnuViewEditor.Checked = False
Exit Sub
End If
If mnuViewEditor.Checked = False Then
frmScript.Show
mnuViewEditor.Checked = True
Exit Sub
End If
End Sub

Private Sub txtInstallColor_Change(Index As Integer)
On Error Resume Next
picInstallColor(Index).BackColor = Val("&H" & txtInstallColor(Index).text)
End Sub

Private Sub txtNSISBGGradientColor_Change(Index As Integer)
On Error Resume Next
picBGColor(Index).BackColor = Val("&H" & txtNSISBGGradientColor(Index).text)
End Sub
