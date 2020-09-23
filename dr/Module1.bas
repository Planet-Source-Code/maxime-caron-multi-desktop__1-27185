Attribute VB_Name = "Module1"
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Function getalltopwindows(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim foregroundwindow As Long
Dim textlen As Long
Dim windowtext As String
Dim svar As Long
Static lastwindowtext As String
foregroundwindow = hWnd
textlen = GetWindowTextLength(foregroundwindow) + 1
windowtext = Space(textlen)
svar = GetWindowText(foregroundwindow, windowtext, textlen)
windowtext = Left(windowtext, Len(windowtext) - 1)
If windowtext = "" Then GoTo slask
If IsWindowVisible(foregroundwindow) > 0 Then
If windowtext = frmMain.Caption Then GoTo slask
frmMain.List1.AddItem windowtext
frmMain.List1.ItemData(frmMain.List1.NewIndex) = foregroundwindow
lastwindowtext = windowtext
End If
slask:
getalltopwindows = 1
End Function




