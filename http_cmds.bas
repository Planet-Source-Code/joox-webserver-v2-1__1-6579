Attribute VB_Name = "http_cmds"
Option Explicit

Global http_port As Long            'Port we will listen on (http port is 80)
Global maxConnections As Long 'Max. # of connections allowed.
Global numConnections As Long  'Number of connections at the time.
Global htmlPageDir As String     'The directory where the
                                               'HTML pages are being stored.
Global iconpath As String
Global html_guestbookstart As String
Global html_guestbookend As String
Global htmlIndexPage As String
Global requestedPage As String

Sub load_defaults()
On Error GoTo Error:
Dim tport As String, i As Integer

'This simply loads up the server.
'Example use:  Call load_defaults

iconpath = "icons\"
http_port = 80 'This is the port we are listening on :)
maxConnections = 100 'Maximium number of connections
'you can have at one time.  After we send the html data,
'the connection is CLOSED.  So, you probably could set this
'to 5 and it would work fine. :)

numConnections = 0 'Total number of connections at the moment is zero :)

htmlIndexPage$ = "index.html"

html_guestbookstart = html_gueststart()
html_guestbookend = html_guestend()

tport$ = ""
If http_port = 80 Then tport$ = "" Else: tport$ = ":" & http_port ' this makes the
                                'string tport ':port'.  the format is http://ip:port.
                                 'if the port is 80 you can just leave it out.(http://ip)

With frmMain
    .sckWS(0).Close
    .sckWS(0).LocalPort = http_port
    .sckWS(0).Listen
    
    .Server.Caption = "http://" & .sckWS(0).LocalIP & tport$ & "/"
End With
ChangeTray "Webserver on - http://" & frmMain.sckWS(0).LocalIP & tport$ & "/", frmMain.ServerOn
frmMain.mnuStart.Caption = "S&top Server"
For i = 1 To maxConnections   'this will preload all winsock controls
    Load frmMain.sckWS(i)
Next i

Exit Sub

Error:
MsgBox "Error on loading winsock.", vbMsgBoxSetForeground + vbInformation
frmMain.sckWS(0).Close
End Sub


Public Sub retrieveHeader(tPage As String, sckWSC)
'This won't be used since this is a web server, not client.
' but i thought i might as well add it. :)

'This is the data sent to a server when the client is requesting
'a page.

'tPage$ is the requested page, e.g., about.html
'sckWSC is a MS winsock control. e.g., Winsock1

'Example use:  Call retrieveHeader(downloads.html, Winsock1)

With sckWSC
    .SendData "GET /" & tPage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: text/plain" & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: " & sckWSC.LocalIP & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf & vbCrLf
End With

End Sub

Public Sub stop_server()
'This sub shuts down the server

With frmMain
.Server.Caption = "Not active"
.sckWS(0).Close 'Closes the port
End With
ChangeTray "Webserver off", frmMain.ServerOff
frmMain.mnuStart.Caption = "&Start Server"

Call unloadControls
Call unset_vars

End Sub

Public Sub unloadControls()
'This unloads all the winsock controls we loaded
Dim i As Integer

With frmMain
For i = 1 To maxConnections
Unload .sckWS(i)
Next i
End With

End Sub


Public Sub unset_vars()
'This clears out all of the varibles

http_port = 0
maxConnections = 0
numConnections = 0
htmlPageDir = 0
htmlIndexPage = ""
iconpath = 0
End Sub


