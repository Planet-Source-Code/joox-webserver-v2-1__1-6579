Attribute VB_Name = "html_data"
'This is all the html data for the webserver. some is imbedded
' into the program, others are external files.
Option Explicit

Public Function html_gueststart()

Dim x As String
x = ""

x = x & "<html>" & vbCrLf
x = x & "<head>" & vbCrLf
x = x & "<style>" & vbCrLf
x = x & "a:link          {font:8pt/11pt verdana; color:red; text-decoration:none}" & vbCrLf
x = x & "a:visited       {font:8pt/11pt verdana; color:red; text-decoration:none}" & vbCrLf
x = x & "a:hover          {font:8pt/11pt verdana; color:red; text-decoration:underline}" & vbCrLf
x = x & "</style>" & vbCrLf
x = x & "<meta HTTP-EQUIV=""Content-Type"" Content=""text-html; charset=Windows-1252"">" & vbCrLf
x = x & "<title>Guestbook</title>" & vbCrLf
x = x & "</head>" & vbCrLf
x = x & "<body bgcolor=""#FFFFFF"">" & vbCrLf
x = x & "<p><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><b>" & vbCrLf
x = x & "<font color=""#FF0000"">Guestbook</font></b></font></p>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
html_gueststart = x
End Function

Public Function html_guestend()

Dim x As String
x = ""

x = x & "<hr>" & vbCrLf
x = x & "<a href=""http://$ip/index.html""><font size=""2""><b>Go back</b></font></a>&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
x = x & "<a href=""http://$ip/addguestbook.html""><font size=""2""><b>Add to guestbook</b></font></a>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
x = x & "<p align=""center""><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"" color=""#999999"">" & vbCrLf
x = x & "  Guestbook-Code, Gui, Counter and more by Joox. &nbsp;E-Mail me at joox@tech-productions.de if you have any questions.</font></p>" & vbCrLf
x = x & "</body>" & vbCrLf
x = x & "</html>" & vbCrLf
html_guestend = x
End Function

Public Function html_filenotfound()
Dim x As String, Name As String, DirBuffer As String, FileBuffer As String, File As Integer
On Error Resume Next

x = x & "<HTML>" & vbCrLf
x = x & "<HEAD>" & vbCrLf
x = x & "<TITLE>Index of " & "/" & ReplaceStr(AddASlash(requestedPage), "\", "/") & "</TITLE>" & vbCrLf
x = x & "</HEAD>" & vbCrLf
x = x & "<BODY>" & vbCrLf
x = x & "<H1>Index of " & "/" & ReplaceStr(AddASlash(requestedPage), "\", "/") & "</H1>" & vbCrLf
x = x & "<PRE><IMG SRC=""" & ReplaceStr("/" & AddASlash(iconpath), "\", "/") & "blank.gif""> Name" & Space(19) & "Size"
x = x & "<HR>" & vbCrLf
x = x & "<IMG SRC=""" & ReplaceStr("/" & AddASlash(iconpath), "\", "/") & "folder.gif""> <A HREF=""./"">..</A>" & Space(21) & "-" & vbCrLf
Name = Dir(AddASlash(htmlPageDir) & AddASlash(requestedPage), vbDirectory)
Do While Name <> ""
   If Name <> "." And Name <> ".." Then
      If (GetAttr(AddASlash(htmlPageDir) & AddASlash(requestedPage) & Name) And vbDirectory) = vbDirectory Then
         DirBuffer = DirBuffer & "<IMG SRC=""" & ReplaceStr("/" & AddASlash(iconpath), "\", "/") & "folder.gif""> <A HREF=""" & ReplaceStr("/" & AddASlash(requestedPage), "\", "/") & Name & """>" & Name & "</A>" & Space(23 - Len(Name)) & "-" & vbCrLf
      Else
         File = FreeFile
         Open AddASlash(htmlPageDir) & AddASlash(requestedPage) & Name For Input As #File
         FileBuffer = FileBuffer & "<IMG SRC=""" & ReplaceStr("/" & AddASlash(iconpath), "\", "/") & "unknown.gif""> <A HREF=""" & ReplaceStr("/" & AddASlash(requestedPage), "\", "/") & Name & """>" & Name & "</A>" & Space(23 - Len(Name)) & LOF(File) & vbCrLf
         Close #File
      End If
   End If
   Name = Dir
   DoEvents 'the system shouldn't freeze
Loop
x = x & DirBuffer & FileBuffer
x = x & "</PRE><HR>" & vbCrLf
x = x & "<ADDRESS>VB-Server 2.1 by Joox at $ip Port " & frmMain.sckWS(0).LocalPort & "</ADDRESS>" & vbCrLf
x = x & "</BODY>" & vbCrLf
x = x & "</HTML>" & vbCrLf
html_filenotfound = x
End Function

