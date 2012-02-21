'请将该部分数据保存为 FORM1.frm 文件
VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "多线程"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6450
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Text            =   "2"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Count"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主线程执行结果测试:"
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   1710
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
'声明了线程ID
    Dim threadid1 As Long
    Dim threadid2 As Long

'参数一，lpThreadAttributes 线程安全属性，传递为NULL 
'参数二，dwStackSize ，线程堆栈大小，可以为0，表示堆栈和此应用堆栈相同
'参数三，lpstartAddress ，执行函数地址，用AddressOf 获取
'参数四，lpParameter ，执行函数的参数地址，可以是一个记录或者是别的类型，用VarPtr获取参数地址（varptr为未公开函数）！！
'参数五，dwCreationFlags ，表示线程创建后的状态！，0表示立即运行，create_SUSPENDED表示线程挂起
'参数六，lpThreadID 表示分配给线程的线程号
    Call CreateThread(Null, ByVal O&, AddressOf Module1.OutText1, VarPtr(0), ByVal 0&, threadid1)
    Call CreateThread(Null, ByVal 0&, AddressOf Module1.OutText2, VarPtr(0), ByVal 0&, threadid2)
    
End Sub

Private Sub Command2_Click()
'该事件运行于主线程！
    Dim i As Long
    i = CLng(Text1.Text)
    Text1.Text = CStr(i * i)  '不要点击次数太多，LONG 类型会溢出
End Sub

Private Sub Form_Load()
'保存窗体句柄全局变量，用于在form 上绘图
    formhandle = Form1.hwnd
End Sub
