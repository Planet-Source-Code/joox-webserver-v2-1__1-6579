VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About Webserver"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3525
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1095
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":058A
      Top             =   360
      Width           =   480
   End
   Begin VB.Label pat 
      Caption         =   "nymainst@nais.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "and"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "Webserver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   795
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Joox 
      Alignment       =   2  'Zentriert
      Caption         =   "joox@tech-productions.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmMain.Enabled = True
AppActivate frmMain.Caption
Unload Me
End Sub

Private Sub Form_Load()
Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
TakeOutMenu Me, SC_CLOSE
End Sub

Private Sub Joox_Click()
Call ShellExecute(Me.hwnd, "Open", "mailto:joox@tech-productions.de", "", "", 1)
End Sub

Private Sub pat_Click()
Call ShellExecute(Me.hwnd, "Open", "mailto:nymainst@nais.com", "", "", 1)
End Sub
