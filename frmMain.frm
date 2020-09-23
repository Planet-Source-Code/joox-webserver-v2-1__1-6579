VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Webserver - Frames Bug Fixed"
   ClientHeight    =   3780
   ClientLeft      =   4875
   ClientTop       =   3825
   ClientWidth     =   3015
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      Begin VB.CheckBox cheCounter 
         Caption         =   "Enable Counter"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox cheActivate 
         Caption         =   "Activate Server on start"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox cheMinimized 
         Caption         =   "Start Minimized"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CheckBox cheGuest 
         Caption         =   "Enable Guestbook"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CheckBox cheLogging 
         Caption         =   "Logging"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdDirChoose 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2540
         TabIndex        =   4
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtRoot 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Webserver is available at:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Server 
         BackColor       =   &H8000000A&
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
         Left            =   120
         TabIndex        =   8
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Server Directory:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3405
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sckWS 
      Index           =   0
      Left            =   240
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image ServerOff 
      Height          =   240
      Left            =   2280
      Picture         =   "frmMain.frx":0E42
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ServerOn 
      Height          =   240
      Left            =   2520
      Picture         =   "frmMain.frx":0F8C
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Server &Options"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStart 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' WebServer code - Frames Bug FIXED - last updated 3/14/2000
'
' E-mail the author at joox@tech-productions.de
'
'This code is an example of how to create
' a advanced webserver.  connect to
' http://your.ip  and it will send the requested
' page. (Requesting "/" will send the index.htm
' page.  if it is not found it sends an index of the directory :)
' you could do plenty with this, i tried to do the most. :)
' if you have any new features which i should add, mail me
'
'There is room for improvement on this, i'm sure.  If you
' improve it, please let me know and send me a copy :)
'
'  Creating a link to another page
' To link to another page, link to 'http://$ip/page_name.html'
'The webserver will replace $ip with your ip. The webserver also
'supports files in other directories in the html dir.  You could link
'to test.html in the directory 'links' by linking to:
' http://$ip/links/test.html
'
'history:
'
'
'        original code was the code from pat(nymainst@nais.com). i added the code to the
'        zip file. if you want to check it out just extract the file OriginalSource.zip.
'
'        i added many, many features i.e. new gui, faster file opening, and and and
'
'latest update:
'  12/26/99   - webserver does now support javacode archives
'  12/26/99   - you can now turn off the counter
'  12/27/99   - the logging function now logs the browser and the requested language of the user
'  03/09/00   - i have finished work!!! the frames-bug is now FIXED!!!!
'  03/12/00   - all winsock controls will be preloaded and will be reused after disconnect!
'               this was a bug in the old code because for every page request a new winsock control
'               was created this has made the code very slow!
'
'
'know bugs:
'             - if you have a network card the ip of her will be displayed in the webserver!
'               even if you were in internet! this will be a problem if you use the $ip variable
'               in you're html file!!
'
'
'                                  PLEASE REPORT ANY BUGS!
'                                to joox@tech-productions.de
'
'
Option Explicit

Private strdata As String


Private Sub cmdDirChoose_Click()
frmDirChoose.Show ownerform:=Me
frmMain.Enabled = False
End Sub

Private Sub cmdOK_Click()
If FileExists(AddASlash(txtRoot.Text)) = False Then
MsgBox "Please enter a valid path for Server Directory.", vbMsgBoxSetForeground + vbInformation
Exit Sub
End If
htmlPageDir = txtRoot.Text
Me.Hide
End Sub

Private Sub Form_Load()

Dim OS As OSVERSIONINFO
OS.dwOSVersionInfoSize = Len(OS)
GetVersionEx OS
If OS.dwMajorVersion < 4 Then
MsgBox "Sorry. You must have Windows 95, Windows 98, NT4 or later!", vbInformation, "Program closed!"
End
End If

If App.PrevInstance Then 'This checks if webserver is allready started
MsgBox "Sorry, but you have Webserver allready started.", vbMsgBoxSetForeground + vbInformation
End
End If

Left = Screen.Width \ 2 - Width \ 2
Top = Screen.Height \ 2 - Height \ 2
TakeOutMenu Me, SC_CLOSE

gHW = Me.hwnd
myNID.cbSize = Len(myNID)
myNID.hwnd = gHW
myNID.uID = uID
myNID.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
myNID.uCallbackMessage = cbNotify
myNID.hIcon = ServerOff
myNID.szTip = "Webserver off" & Chr(0)
ShellNotifyIcon NIM_ADD, myNID
Hook
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3

Server.Caption = "Not active"

If FileExists(AddASlash(App.Path) & "Webserver.ini") = True Then
Dim Cache As String, Files As Integer
Files = FreeFile
Open AddASlash(App.Path) & "Webserver.ini" For Input As #Files
Do While Not EOF(Files)
Line Input #Files, Cache
If Mid(Cache, 1, 1) <> "[" Then
If Mid(Cache, 1, 10) = "ServerRoot" Then
If FileExists(AddASlash(Mid(Cache, 12, Len(Cache)))) = True Then
txtRoot.Text = Mid(Cache, 12, Len(Cache))
Else
txtRoot.Text = App.Path
End If
ElseIf Mid(Cache, 1, 7) = "Logging" Then
If Mid(Cache, 9, 1) = "1" Then
cheLogging.Value = 1
End If
ElseIf Mid(Cache, 1, 9) = "Guestbook" Then
If Mid(Cache, 11, 1) = "1" Then
cheGuest.Value = 1
End If
ElseIf Mid(Cache, 1, 7) = "Counter" Then
If Mid(Cache, 9, 1) = "1" Then
cheCounter.Value = 1
End If
ElseIf Mid(Cache, 1, 9) = "Minimized" Then
If Mid(Cache, 11, 1) = "1" Then
cheMinimized = 1
Me.Hide
End If
ElseIf Mid(Cache, 1, 15) = "ActivateOnStart" Then
If Mid(Cache, 17, 1) = "1" Then
cheActivate.Value = 1
load_defaults
End If
End If
End If
Loop
Close #Files
Else
If FileExists(AddASlash(App.Path) & "html\") Then
txtRoot.Text = AddASlash(App.Path) & "html\"
Else
txtRoot.Text = App.Path
End If
cheGuest.Value = 1
cheCounter.Value = 1
cheLogging.Value = 1
cheMinimized.Value = 1
cheActivate.Value = 0
End If
htmlPageDir = txtRoot.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Files As String, Buffer As String
Call stop_server
Files = FreeFile
Open AddASlash(App.Path) & "Webserver.ini" For Output As Files
Buffer = ""
Buffer = "[Webserver Options]" & vbCrLf
Buffer = Buffer & "ServerRoot=" & txtRoot.Text & vbCrLf
Buffer = Buffer & "Logging=" & cheLogging.Value & vbCrLf
Buffer = Buffer & "Guestbook=" & cheGuest.Value & vbCrLf
Buffer = Buffer & "Counter=" & cheCounter.Value & vbCrLf
Buffer = Buffer & "Minimized=" & cheMinimized & vbCrLf
Buffer = Buffer & "ActivateOnStart=" & cheActivate.Value & vbCrLf
Print #Files, Buffer
Close #Files
SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 3
Unhook
ShellNotifyIcon NIM_DELETE, myNID
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show ownerform:=Me
frmMain.Enabled = False
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOptions_Click()
frmMain.Visible = True
AppActivate frmMain.Caption
End Sub

Private Sub mnuStart_Click()
If mnuStart.Caption = "&Start Server" Then
load_defaults
Else
stop_server
End If
End Sub

Private Sub sckWS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
   If Index = 0 Then
      numConnections = numConnections + 1 'number of connected clients + 1
    
      If numConnections = maxConnections Then GoTo done 'if we've reached the max # of connections, exit sub.
      Dim i As Integer
      For i = 1 To maxConnections 'this for...next will search for a free socket
         If sckWS(i).State = sckClosed Then
            Exit For
         End If
      Next i
      sckWS(i).LocalPort = 0 'set its local port to 0
      sckWS(i).Accept requestID 'Accept the connection request.
      Exit Sub
      
done:
      numConnections = numConnections - 1
      
End If
End Sub

Private Sub sckWS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
Dim spc2 As Integer
sckWS(Index).GetData strdata$ 'Get any data sent to us
'MsgBox strdata$ ' I used this for debugging
If Mid(strdata$, 1, 3) = "GET" Then 'If it is trying to get a site, find out
Dim findget As String, pagetoget As String
findget = InStr(strdata, "GET ")      ' the site they want then set requestedPage$
spc2 = InStr(findget + 5, strdata, " ") ' to it.
pagetoget = Mid(strdata, findget + 4, spc2 - (findget + 4))
requestedPage = pagetoget
ElseIf Mid$(strdata, 1, 4) = "POST" Then 'This is the code when it is trying to post something!
Dim findpost As String, pagetopost As String
findpost = InStr(strdata$, "POST ")        'the data where filtered in the ConnectionRequest
spc2 = InStr(findpost + 5, strdata, " ")   'Function of the winsock control
pagetopost = Mid(strdata, findpost + 5, spc2 - (findpost + 5))
requestedPage = pagetopost
End If

If Left(requestedPage, Len(iconpath) + 1) = "/" & ReplaceStr(iconpath, "\", "/") Then 'this will check if it is a icon
sckWS(Index).SendData text_read(AddASlash(App.Path) & requestedPage)
Exit Sub
End If

      If cheLogging.Value = 1 Then
      Dim Logging As Integer
      Logging = FreeFile      'This is for the logging function
      Open AddASlash(App.Path) & "Log.log" For Append As #Logging
      Print #Logging, Format(Date, "Long Date") & " " & Format(Time, "Long Time") & " ; " & sckWS(Index).RemoteHostIP & "; " & Mid(strdata$, InStr(1, UCase(strdata$), "USER-AGENT:") + 12, InStr(InStr(1, UCase(strdata$), "USER-AGENT:") + 12, UCase(strdata$), vbCrLf) - InStr(1, UCase(strdata$), "USER-AGENT:") - 12) & "; requested Language: " & Mid(strdata$, InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") + 17, InStr(InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") + 17, UCase(strdata$), vbCrLf) - InStr(1, UCase(strdata$), "ACCEPT-LANGUAGE:") - 17) & "; requested page: " & requestedPage$
      Close #Logging
      End If
      
      If requestedPage$ = "/" Then
      requestedPage$ = htmlIndexPage$ ' if the page '/' was requested, set requested page to the index html page.
      Else
      requestedPage$ = Mid(requestedPage$, 2, Len(requestedPage$) - 1)
      End If
      
      If cheGuest.Value = 1 Then
      Dim NameStart As Integer, NameEnd As Integer, NameValue As String, MailStart As Integer, MailEnd As Integer, MailValue As String, CommentStart As Integer, CommentEnd As Integer, CommentValue As String, Guestbook As Integer, datastr As String
      If UCase(requestedPage$) = "GUESTBOOK.CGI" Then 'This is check if the Guestbook.cgi is requested
      NameStart = InStr(UCase(strdata$), "NAME=")
      NameEnd = InStr(NameStart + 5, strdata$, "&")
      NameValue = Mid$(strdata$, NameStart + 5, NameEnd - (NameStart + 5))
      MailStart = InStr(UCase(strdata$), "E-MAIL=")
      MailEnd = InStr(MailStart + 7, strdata$, "&")
      MailValue = Mid$(strdata$, MailStart + 7, MailEnd - (MailStart + 7))
      CommentStart = InStr(UCase(strdata$), "COMMENT=")
      CommentEnd = InStr(CommentStart + 8, strdata$, "&")
      CommentValue = Mid$(strdata$, CommentStart + 8, CommentEnd - (CommentStart + 8))
      CommentValue = ReplaceStr(CommentValue, "+", " ")
      CommentValue = ReplaceStr(CommentValue, "%0D%0A", "<br>")
      CommentValue = ReplaceStr(CommentValue, "%21", "!")
      CommentValue = ReplaceStr(CommentValue, "%22", "&quot;")
      CommentValue = ReplaceStr(CommentValue, "%A7", "§")
      CommentValue = ReplaceStr(CommentValue, "%24", "$")
      CommentValue = ReplaceStr(CommentValue, "%25", "%")
      CommentValue = ReplaceStr(CommentValue, "%26", "&")
      CommentValue = ReplaceStr(CommentValue, "%2F", "/")
      CommentValue = ReplaceStr(CommentValue, "%28", "(")
      CommentValue = ReplaceStr(CommentValue, "%29", ")")
      CommentValue = ReplaceStr(CommentValue, "%3D", "=")
      CommentValue = ReplaceStr(CommentValue, "%3F", "?")
      CommentValue = ReplaceStr(CommentValue, "%B2", "²")
      CommentValue = ReplaceStr(CommentValue, "%B3", "³")
      CommentValue = ReplaceStr(CommentValue, "%7B", "{")
      CommentValue = ReplaceStr(CommentValue, "%5B", "[")
      CommentValue = ReplaceStr(CommentValue, "%5D", "]")
      CommentValue = ReplaceStr(CommentValue, "%7D", "}")
      CommentValue = ReplaceStr(CommentValue, "%5C", "\")
      CommentValue = ReplaceStr(CommentValue, "%DF", "ß")
      CommentValue = ReplaceStr(CommentValue, "%23", "#")
      CommentValue = ReplaceStr(CommentValue, "%27", "'")
      CommentValue = ReplaceStr(CommentValue, "%3A", ":")
      CommentValue = ReplaceStr(CommentValue, "%2C", ",")
      CommentValue = ReplaceStr(CommentValue, "%3B", ";")
      CommentValue = ReplaceStr(CommentValue, "%60", "`")
      CommentValue = ReplaceStr(CommentValue, "%7E", "~")
      CommentValue = ReplaceStr(CommentValue, "%2B", "+")
      CommentValue = ReplaceStr(CommentValue, "%B4", "´")
      MailValue = ReplaceStr(MailValue, "%21", "!")
      MailValue = ReplaceStr(MailValue, "%22", "&quot;")
      MailValue = ReplaceStr(MailValue, "%A7", "§")
      MailValue = ReplaceStr(MailValue, "%24", "$")
      MailValue = ReplaceStr(MailValue, "%25", "%")
      MailValue = ReplaceStr(MailValue, "%26", "&")
      MailValue = ReplaceStr(MailValue, "%2F", "/")
      MailValue = ReplaceStr(MailValue, "%28", "(")
      MailValue = ReplaceStr(MailValue, "%29", ")")
      MailValue = ReplaceStr(MailValue, "%3D", "=")
      MailValue = ReplaceStr(MailValue, "%3F", "?")
      MailValue = ReplaceStr(MailValue, "%B2", "²")
      MailValue = ReplaceStr(MailValue, "%B3", "³")
      MailValue = ReplaceStr(MailValue, "%7B", "{")
      MailValue = ReplaceStr(MailValue, "%5B", "[")
      MailValue = ReplaceStr(MailValue, "%5D", "]")
      MailValue = ReplaceStr(MailValue, "%7D", "}")
      MailValue = ReplaceStr(MailValue, "%5C", "\")
      MailValue = ReplaceStr(MailValue, "%DF", "ß")
      MailValue = ReplaceStr(MailValue, "%23", "#")
      MailValue = ReplaceStr(MailValue, "%27", "'")
      MailValue = ReplaceStr(MailValue, "%3A", ":")
      MailValue = ReplaceStr(MailValue, "%2C", ",")
      MailValue = ReplaceStr(MailValue, "%3B", ";")
      MailValue = ReplaceStr(MailValue, "%60", "`")
      MailValue = ReplaceStr(MailValue, "%7E", "~")
      MailValue = ReplaceStr(MailValue, "%2B", "+")
      MailValue = ReplaceStr(MailValue, "%B4", "´")
      NameValue = ReplaceStr(NameValue, "%21", "!")
      NameValue = ReplaceStr(NameValue, "%22", "&quot;")
      NameValue = ReplaceStr(NameValue, "%A7", "§")
      NameValue = ReplaceStr(NameValue, "%24", "$")
      NameValue = ReplaceStr(NameValue, "%25", "%")
      NameValue = ReplaceStr(NameValue, "%26", "&")
      NameValue = ReplaceStr(NameValue, "%2F", "/")
      NameValue = ReplaceStr(NameValue, "%28", "(")
      NameValue = ReplaceStr(NameValue, "%29", ")")
      NameValue = ReplaceStr(NameValue, "%3D", "=")
      NameValue = ReplaceStr(NameValue, "%3F", "?")
      NameValue = ReplaceStr(NameValue, "%B2", "²")
      NameValue = ReplaceStr(NameValue, "%B3", "³")
      NameValue = ReplaceStr(NameValue, "%7B", "{")
      NameValue = ReplaceStr(NameValue, "%5B", "[")
      NameValue = ReplaceStr(NameValue, "%5D", "]")
      NameValue = ReplaceStr(NameValue, "%7D", "}")
      NameValue = ReplaceStr(NameValue, "%5C", "\")
      NameValue = ReplaceStr(NameValue, "%DF", "ß")
      NameValue = ReplaceStr(NameValue, "%23", "#")
      NameValue = ReplaceStr(NameValue, "%27", "'")
      NameValue = ReplaceStr(NameValue, "%3A", ":")
      NameValue = ReplaceStr(NameValue, "%2C", ",")
      NameValue = ReplaceStr(NameValue, "%3B", ";")
      NameValue = ReplaceStr(NameValue, "%60", "`")
      NameValue = ReplaceStr(NameValue, "%7E", "~")
      NameValue = ReplaceStr(NameValue, "%2B", "+")
      NameValue = ReplaceStr(NameValue, "%B4", "´")
      NameValue = ReplaceStr(NameValue, "+", " ")
      Guestbook = FreeFile
      Open AddASlash(App.Path) & "guestbook.ini" For Append As #Guestbook
      datastr = "<b><u>Name:</u></b>&nbsp;&nbsp;" & NameValue
      datastr = datastr & "&nbsp;&nbsp;&nbsp;<b><u>E-Mail:</u></b>&nbsp;&nbsp;<a href=mailto:" & MailValue
      datastr = datastr & ">" & MailValue
      datastr = datastr & "</a><br><br><b><u>Comment:</u></b><br>" & CommentValue
      datastr = datastr & "<br><br><br><br>"
      Print #Guestbook, datastr
      Close #Guestbook
      strdata$ = ""
      requestedPage$ = "guestbook.html"
      End If
      
      If UCase(requestedPage$) = "GUESTBOOK.HTML" Then
      Dim htmldata As String
      htmldata = html_guestbookstart & vbCrLf & text_read(AddASlash(App.Path) & "guestbook.ini") & vbCrLf & html_guestbookend & vbCrLf
      sckWS(Index).SendData ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP)
      GoTo done
      End If
      End If
      
      If FileExists(AddASlash(htmlPageDir) & requestedPage$) Then 'if the requested page exists, then..
      
      htmldata$ = text_read(AddASlash(htmlPageDir) & requestedPage$) 'This reads the file and stores it's contents in htmldata$
      
      If cheCounter.Value = 1 Then
      Dim CountValue As String, Counter As Integer
      If InStr(1, htmldata$, "$counter") <> 0 Then 'Checks if $counter is in the html page
      If FileExists(AddASlash(App.Path) & "counter.ini") Then  ' if true the counter will count one up
      CountValue = text_read(AddASlash(App.Path) & "counter.ini")
      Else
      CountValue = "0"
      End If
      CountValue = CountValue + 1
      Counter = FreeFile
      Open AddASlash(App.Path) & "counter.ini" For Output As #Counter
      Print #Counter, CountValue
      Close #Counter
      htmldata$ = ReplaceStr(htmldata$, "$counter", Str(CountValue))
      End If
      End If
      
      htmldata$ = ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP) 'Oops, i didn't use the replace function right.  Now it's fixed at replaces $ip with your IP.
      sckWS(Index).SendData htmldata$ & vbCrLf
      
      ElseIf FileExists(AddASlash(htmlPageDir) & AddASlash(requestedPage$) & htmlIndexPage$) Then
      htmldata$ = text_read(AddASlash(htmlPageDir) & AddASlash(requestedPage$) & htmlIndexPage$)
      htmldata$ = ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP)
      sckWS(Index).SendData htmldata$ & vbCrLf

      Else 'if the file doesn't exists
      
      If requestedPage = htmlIndexPage Then requestedPage = ""
      sckWS(Index).SendData ReplaceStr(html_filenotfound, "$ip", sckWS(0).LocalIP) & vbCrLf
      End If

done:
numConnections = numConnections - 1 'number of connections at the moment - 1
      
End Sub


Private Sub sckWS_SendComplete(Index As Integer)
      
      sckWS(Index).Close 'Close the connection.

End Sub


Private Sub Server_Click()
If Mid(Server.Caption, 1, 7) = "http://" Then
Call ShellExecute(Me.hwnd, "Open", Server.Caption, "", "", 1)
End If
End Sub

