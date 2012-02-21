VERSION 5.00
Begin VB.Form frmSetDB 
   Caption         =   "设置数据库路径"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4125
   StartUpPosition =   1  '所有者中心
   Begin DataRepair.XPButton XPButton2 
      Height          =   360
      Left            =   540
      TabIndex        =   3
      Top             =   1005
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      Caption         =   "恢复默认值"
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
   Begin DataRepair.XPButton XPButton1 
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Top             =   1005
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   635
      Caption         =   "确定"
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
   Begin DataRepair.FTextBox FTextBox1 
      Height          =   300
      Left            =   1485
      TabIndex        =   0
      Top             =   390
      Width           =   2145
      _ExtentX        =   3784
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
   End
   Begin VB.Label Label1 
      Caption         =   "数据库路径"
      Height          =   390
      Left            =   390
      TabIndex        =   2
      Top             =   450
      Width           =   930
   End
End
Attribute VB_Name = "frmSetDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag As Boolean
Private Sub Form_Load()
    FTextBox1.Text = dbPath
    flag = False
    
End Sub

Private Sub FTextBox1_Change()
    flag = True
End Sub

Private Sub XPButton1_Click()
    If flag Then
        dbPath = FTextBox1.Text
        Open GetApp & "conf.txt" For Output As #1
        Print #1, FTextBox1.Text
        Close #1
    End If
    Unload Me
End Sub

Private Sub XPButton2_Click()
    FTextBox1.Text = "e:\smis\data\smis.mdb"
End Sub
