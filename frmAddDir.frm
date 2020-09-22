VERSION 5.00
Begin VB.Form frmAddDir 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select a directory"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   4455
   End
   Begin VB.DirListBox dirAdd 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   505
      Width           =   4455
   End
   Begin VB.DriveListBox drvAdd 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.FileListBox fileAdd 
      Height          =   2430
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmAddDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
'this sub recurses the subdirectories & files
On Error Resume Next
'recurse
frmMain.DoDirs dirAdd.path, "*.*"
'hide this form
Me.Hide
End Sub

Private Sub dirAdd_Change()
'this sub changes the filelist
On Error Resume Next
fileAdd.path = dirAdd.path
End Sub

Private Sub drvAdd_Change()
'this sub changes the directory list
On Error Resume Next
dirAdd.path = drvAdd.Drive
End Sub
