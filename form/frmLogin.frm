VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "祥兴达仓管系统--数据修复"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  '屏幕中心
   Begin DataRepair.XPButton SetDB 
      Height          =   375
      Left            =   300
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "数据库设置"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin DataRepair.FCombo cboUser 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnabledText     =   0   'False
      ListIndex       =   -1
   End
   Begin DataRepair.FTextBox txtPW 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "宋体"
      FontSize        =   9
      PasswordChar    =   "*"
   End
   Begin DataRepair.XPButton cmdOK 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "登录(&L)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin DataRepair.XPButton cmdExit 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "退出(&Q)"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   5250
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密  码："
      Height          =   180
      Left            =   600
      TabIndex        =   1
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   1500
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFF8F0&
      BorderColor     =   &H00C5742F&
      Height          =   1335
      Left            =   300
      Top             =   1140
      Width           =   3900
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboUser_GotFocus()
    cboUser.SelAll
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
'On Error GoTo aaaa
    Dim rs As New adodb.Recordset, strMD5 As String
    If Conn.State <> 0 Then Conn.Close
    DBConnect
    rs.Open "Select * From UserInfo Where User='" & cboUser.Text & "'", Conn, 1, 1
    
    Do While Not rs.EOF
        If Trim(txtPW.Text) = jiemi(rs("PWD")) Then
            rs.Close
            Set rs = Nothing
            Conn.Close
            Set Conn = Nothing
            Unload Me
            frmMain.Show
            Exit Sub
        End If
        rs.MoveNext
    Loop
    
    MsgBox "用户名或密码错误，登陆失败！", vbCritical, "登录"
    rs.Close
    Conn.Close
Exit Sub
aaaa:
    MsgBox Err.Description, vbCritical
    If Conn.State = 1 Then Conn.Close
End Sub
Private Function jiemi(strf)
   sz = Asc(Right(strf, 1))
   
   tmp = ""
   
   For i = 1 To Len(strf) - 1
       tmp = tmp & Chr((Asc(Mid(strf, i, 1)) Xor sz))
   Next
   
   jiemi = tmp

End Function
Private Sub cmdServer_Click()
    With frmServer
        .txtServer.Text = strSQLServer
        .txtUser.Text = strSQLUser
        If strSQLPW <> "" Then .lbPW.Visible = True
        .txtDB.Text = IIf(strSQLDB <> "", strSQLDB, "SuperMarketdb")
        .Show 1
    End With
End Sub

Private Sub Form_Activate()
On Error Resume Next
    cboUser.SetFocus
    cboUser.SetF
    If Conn.State <> 0 Then Conn.Close
    LoadUserList
 
    If cboUser.ListCount > 0 Then cboUser.ListIndex = 0
    
    txtPW.SetFocus
End Sub

Public Sub LoadUserList()
On Error GoTo ErrProcess
    Dim rs As adodb.Recordset
    Dim strSQL As String
    
    Set rs = New adodb.Recordset
    
    strSQL = "select * from userInfo"
    DBConnect
    rs.Open strSQL, Conn, 1, 1
    
    If rs.EOF Then
        MsgBox "未找到用户，程序将关闭", vbCritical, "登录"
        
    Else
        Do Until rs.EOF
            cboUser.AddItem Trim(rs("user"))
            rs.MoveNext
        Loop
    End If

    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing

    Exit Sub
    
ErrProcess:
    MsgBox Err.Description, vbInformation, "登录"
    
End Sub
Private Sub Form_Load()
On Error GoTo errPro
    dbPath = ""
    Open GetApp & "conf.txt" For Input As #1
    Line Input #1, dbPath
    Close #1
    If dbPath = "" Then dbPath = "f:\smis\data\smis.mdb"
    Exit Sub
errPro:
    dbPath = "f:\smis\data\smis.mdb"
End Sub
Private Sub SetDB_Click()
    frmSetDB.Show vbModal, frmLogin

End Sub
Private Sub txtPW_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdOK_Click
    End If
End Sub
