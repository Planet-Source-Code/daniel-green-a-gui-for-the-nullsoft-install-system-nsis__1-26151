VERSION 5.00
Begin VB.Form frmScript 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Script Editor"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Script"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtParse 
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.TextBox txtNSISScript 
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
frmMain.mnuViewEditor.Checked = False
End Sub
