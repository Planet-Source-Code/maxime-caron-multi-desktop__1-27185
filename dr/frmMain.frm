VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "virtual desktop by skyde"
   ClientHeight    =   6825
   ClientLeft      =   285
   ClientTop       =   -915
   ClientWidth     =   7365
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3850.494
   ScaleMode       =   0  'User
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   125
      Left            =   480
      Top             =   2160
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   6825
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7410
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuStartAnimation 
         Caption         =   "Start &Animation"
      End
      Begin VB.Menu mnuStopAnimation 
         Caption         =   "&Stop Animation"
      End
      Begin VB.Menu SEP01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Window"
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
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private WithEvents SysTray As CSysTray
Attribute SysTray.VB_VarHelpID = -1
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Dim d As Boolean
Private Const VK_F9 = &H78
Private Sub patate()



List1.Clear
svar = EnumWindows(AddressOf getalltopwindows, 0)


    Set SysTray = New CSysTray
    Set SysTray.SourceWindow = Me
    
    SysTray.ChangeIcon App.Path & "\globe.ani"
    SysTray.ToolTip = Me.Caption
    
    SysTray.MinToSysTray

    
End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        SysTray.MinToSysTray
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    

    SysTray.RemoveFromSysTray

End Sub



Private Sub Image1_Click()
patate
End Sub

Private Sub SysTray_lButtonUP()

If d Then

   d = False

List2.Clear
svar = EnumWindows(AddressOf getalltopwindows2, 0)
For i = 0 To List2.ListCount - 1
a = List2.ItemData(i)
a = ShowWindow(a, 0) ' hide
Next
For i = 0 To List1.ListCount - 1
a = List1.ItemData(i)
a = ShowWindow(a, 5) ' show
Next
Else
    d = True
 
    List1.Clear
    svar = EnumWindows(AddressOf getalltopwindows, 0)
    For i = 0 To List1.ListCount - 1
        If List1.List(i) <> "Program Manager" Then
            a = List1.ItemData(i)
            a = ShowWindow(a, 0) ' hide
        End If
    Next
    For i = 0 To List2.ListCount - 1
        a = List2.ItemData(i)
        a = ShowWindow(a, 5) ' show
    Next
End If

End Sub
Private Sub SysTray_rButtonUP()

For i = 0 To List2.ListCount - 1
a = List2.ItemData(i)
a = ShowWindow(a, 5) ' show
Next
For i = 0 To List1.ListCount - 1
a = List1.ItemData(i)
a = ShowWindow(a, 5) ' show
Next
End
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
    patate
End Sub







Private Sub Timer2_Timer()
If GetAsyncKeyState(VK_F9) Then
SysTray_lButtonUP
End If
End Sub
