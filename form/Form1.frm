'�뽫�ò������ݱ���Ϊ FORM1.frm �ļ�
VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���߳�"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6450
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   960
      TabIndex        =   2
      Text            =   "2"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
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
      Caption         =   "���߳�ִ�н������:"
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
'�������߳�ID
    Dim threadid1 As Long
    Dim threadid2 As Long

'����һ��lpThreadAttributes �̰߳�ȫ���ԣ�����ΪNULL 
'��������dwStackSize ���̶߳�ջ��С������Ϊ0����ʾ��ջ�ʹ�Ӧ�ö�ջ��ͬ
'��������lpstartAddress ��ִ�к�����ַ����AddressOf ��ȡ
'�����ģ�lpParameter ��ִ�к����Ĳ�����ַ��������һ����¼�����Ǳ�����ͣ���VarPtr��ȡ������ַ��varptrΪδ��������������
'�����壬dwCreationFlags ����ʾ�̴߳������״̬����0��ʾ�������У�create_SUSPENDED��ʾ�̹߳���
'��������lpThreadID ��ʾ������̵߳��̺߳�
    Call CreateThread(Null, ByVal O&, AddressOf Module1.OutText1, VarPtr(0), ByVal 0&, threadid1)
    Call CreateThread(Null, ByVal 0&, AddressOf Module1.OutText2, VarPtr(0), ByVal 0&, threadid2)
    
End Sub

Private Sub Command2_Click()
'���¼����������̣߳�
    Dim i As Long
    i = CLng(Text1.Text)
    Text1.Text = CStr(i * i)  '��Ҫ�������̫�࣬LONG ���ͻ����
End Sub

Private Sub Form_Load()
'���洰����ȫ�ֱ�����������form �ϻ�ͼ
    formhandle = Form1.hwnd
End Sub
