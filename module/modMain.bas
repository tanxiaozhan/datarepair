Attribute VB_Name = "modMain"
Public GetApp As String
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_SHOWWINDOW = &H40
Dim retValue As Long
Public isTop As Boolean



'程序入口
Public Sub Main()
'On Error Resume Next
    
    If App.PrevInstance Then
        End
        Exit Sub
    End If
    '获得本地路径
    '获得本地路径
    GetApp = App.Path: If Right$(GetApp, 1) <> "\" Then GetApp = GetApp & "\"
    isTop = True
    'frmMain.Show
    frmLogin.Show
End Sub
