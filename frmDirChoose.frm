VERSION 5.00
Begin VB.Form frmDirChoose 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Choose Directory"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2930
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   1
      Top             =   330
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmDirChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
frmMain.txtRoot.Text = Dir1.Path
frmMain.Enabled = True
AppActivate frmMain.Caption
Unload Me
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
On Error Resume Next
Drive1.Drive = Mid(frmMain.txtRoot.Text, 1, 2)
Dir1.Path = frmMain.txtRoot.Text
TakeOutMenu Me, SC_CLOSE
Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub
