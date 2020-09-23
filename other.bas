Attribute VB_Name = "other"
'Other functions and subs needed :)

Public Function ReplaceStr(ByVal strMain As String, strFind As String, strReplace As String) As String
'Thsi is the same thing as the Replace function in vb6.  I added this
'for those of you using vb5.  This was NOT written by me, it was written by
' someone named 'dos'.  He's a great programmer, visit his webpage @
' http://hider.com/dos

    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    ReplaceStr$ = strNew$
End Function

Public Function PathFromString(fullString As String)
'Thanks to Johannes Eder for this function

    fullString = ReplaceStr(fullString, "/", "\")
    For i = 1 To Len(fullString)
    buff$ = Mid$(fullString, Len(fullString) - (i - 1), 1)
    If buff$ = "\" Then
    fullString = Left$(fullString, Len(fullString) - (i - 1)): GoTo done
    End If
    Next i
done:
    If Left$(fullString, 1) = "\" Then
    fullString = Mid$(fullString, 2, Len(fullString) - 1)
    End If
    PathFromString = fullString
    
End Function

Public Function text_read(filename)
'This function reads a file and spits out the text in it.

Dim f
Dim textda
Dim cha

On Error Resume Next

i = 1
f = FreeFile
textda = ""
If FileExists(filename) Then
    If Len(filename) Then
        Open filename For Binary As #f   ' Open file.
            textda = Input(LOF(f), #f) ' I HAVE CHANGED THIS FROM 1 TO LOF(f) BECAUSE OF BIG FILES (100 KB)
           ' DoEvents  'I HAVE ADDED THIS FOR BIG FILES
        Close #f
    End If
text_read = textda
Else
text_read = ""
End If

End Function
Public Function FileExists(ByVal sFileName As String) As Integer
'Checks if the given file exists.

Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        FileExists = False
        Else
            FileExists = True
    End If
End Function
Public Sub timeout(ByVal nSecond As Single)
'Pauses for x seconds.

   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim Dummy As Integer

      Dummy = DoEvents()
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub

Public Function ConvertString(tmpVal As String, KeyValSize As Long) As String
   If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        ConvertString = Left(tmpVal, KeyValSize - 1)
    Else
        ConvertString = Left(tmpVal, KeyValSize)
    End If
End Function

Public Function RegReadValue(Stamm As Long, Pfad As String, Schluessel As String) As String
Dim dataBuff As String, ldataBuffSize As Long, phkResult As Long, retval As Long, Text As String
    dataBuff = Space(255)
    ldataBuffSize = Len(dataBuff)
    retval = RegOpenKeyEx(Stamm, Pfad, 0, KEY_ALL_ACCESS, phkResult)
    retval = RegQueryValueEx(phkResult, Schluessel, 0, 0, dataBuff, ldataBuffSize)
    If retval = ERROR_SUCCESS Then
            RegReadValue = ConvertString(dataBuff, ldataBuffSize)
    Else
            RegReadValue = "Error"
    End If
    RegCloseKey Stamm
    RegCloseKey phkResult
End Function

Public Function RegWriteKey(Stamm As Long, Pfad As String) As Long
Dim retval As Long, phkResult As Long, SA As SECURITY_ATTRIBUTES, Create As Long
retval = RegCreateKeyEx(Stamm, Pfad, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, Create)
RegCloseKey phkResult
End Function

Public Function RegDelKey(Stamm As Long, Pfad As String) As Long
Dim retval As Long, phkResult As Long
retval = RegDeleteKey(Stamm, Pfad)
RegCloseKey phkResult
End Function

Public Function RegDelValue(Stamm As Long, Pfad As String, Value As String) As Long
Dim retval As Long, phkResult As Long
Pfad = AddASlash(Pfad)
retval = RegOpenKeyEx(Stamm, Pfad, 0, KEY_ALL_ACCESS, phkResult)
retval = RegDeleteValue(Stamm, Value)
RegCloseKey phkResult
End Function

Public Function RegWriteValue(Stamm As Long, Pfad As String, Value As String, Wert As String) As Long
Dim retval As Long, phkResult As Long, SA As SECURITY_ATTRIBUTES, Create As Long
Pfad = AddASlash(Pfad)
retval = RegCreateKeyEx(Stamm, Pfad, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, phkResult, Create)
    retval = RegSetValueEx(phkResult, Value, 0, REG_SZ, _
        Wert, CLng(Len(Wert) + 1))
    RegCloseKey phkResult
End Function

Public Function AddASlash(InString As String) As String
  If InString <> "" Then
    If Mid(InString, Len(InString), 1) <> "\" Then
        AddASlash = InString & "\"
    Else
        AddASlash = InString
    End If
  End If
End Function

Public Sub Hook()
lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
Dim tmp As Long
tmp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
If wParam = uID Then
Select Case lParam
Case WM_MOUSEMOVE
Case WM_LBUTTONDOWN
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK

If frmMain.mnuStart.Caption = "&Start Server" Then
load_defaults
Else
stop_server
End If

'frmMain.Visible = True
'AppActivate frmMain.Caption
Case WM_RBUTTONDOWN

frmMain.PopupMenu frmMain.mnuTray, vbPopupMenuRightAlign, , , frmMain.mnuStart

Case WM_RBUTTONUP
Case WM_RBUTTONDBLCLK
Case WM_MBUTTONDOWN
Case WM_MBUTTONUP
Case WM_MBUTTONDBLCLK
Case Else
End Select
End If
WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Public Sub ChangeTray(Title As String, Icon As Object)
myNID.hIcon = Icon
myNID.szTip = Title & Chr(0)
ShellNotifyIcon NIM_MODIFY, myNID
End Sub

Public Function TakeOutMenu(ThisForm As Form, ParamArray MenusToRemove() As Variant)
    Dim DeleteMenu As Long
    Dim ControlMenuHwnd As Long
    Dim RemoveItem As Integer
    Dim HighestArrayNumber
    Dim x As Integer
    
    HighestArrayNumber = Val(UBound(MenusToRemove))
      
    For x = 0 To 5
        'If no parameters were passed, then just exit
        If HighestArrayNumber = -1 Then
            MsgBox "No parameters specified"
            Exit Function
        End If
        'If 6 or less arguments are passed, then
        'we must exit when we get to the last element
        'of the list!
        If x > HighestArrayNumber Then
           Exit Function
        End If
        'Take out the specified menu item now
        RemoveItem = Val(MenusToRemove(x))
        'Retrieve the Control Menu's handle
        ControlMenuHwnd = GetSystemMenu(ThisForm.hwnd, 0)
        'Remove this menu item
        DeleteMenu = RemoveMenu(ControlMenuHwnd, RemoveItem, MF_BYCOMMAND)
    Next x
End Function
