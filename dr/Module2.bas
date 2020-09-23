Attribute VB_Name = "Module2"
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Function getalltopwindows2(ByVal hWnd2 As Long, ByVal lParam2 As Long) As Long
Dim foregroundwindow2 As Long
Dim textlen2 As Long
Dim windowtext2 As String
Dim svar2 As Long
Static lastwindowtext2 As String
foregroundwindow2 = hWnd2
textlen2 = GetWindowTextLength(foregroundwindow2) + 1
windowtext2 = Space(textlen2)
svar2 = GetWindowText(foregroundwindow2, windowtext2, textlen2)
windowtext2 = Left(windowtext2, Len(windowtext2) - 1)
If windowtext2 = "" Then GoTo slask
If IsWindowVisible(foregroundwindow2) > 0 Then
If windowtext2 = frmMain.Caption Then GoTo slask
frmMain.List2.AddItem windowtext2
frmMain.List2.ItemData(frmMain.List2.NewIndex) = foregroundwindow2
lastwindowtext2 = windowtext2
End If
slask:
getalltopwindows2 = 1
End Function
