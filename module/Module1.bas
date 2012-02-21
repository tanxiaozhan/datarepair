Attribute VB_Name = "Module1"
'请将该部分数据保存为 Module1.bas 文件

'线程安全属性数据结构；
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'这个是用于多线程访问临界资源同步Api的数据结构
Public Type CRITICAL_SECTION
    dummy As Long
End Type
'为什么用GDI 函数绘图？原因等下再讲
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'请注意；createThread APi声明已被我修改过，修改的地方请自行参照APIView复制的内容
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
'这个是sleep,作用就是让两个线程绘图频率不一致，效果才明显。
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '进入临界区
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '离开临界区

'几个重要的函数举例
'ObjPtr：返回对象实例私有域的地址。
'StrPtr：返回字符串第一个字的地址。
'VarPtr：返回变量的地址。

'全局的form的句柄！
Public formhandle As Long
'临界数据结构
Public sect As CRITICAL_SECTION
Public pro As Long

Sub OutText1()  '过程一
Dim i As Long
Dim dc As Long
Dim s As String
    dc = GetDC(formhandle) '获取窗体句柄的DC
        Call SetBkColor(dc, &HF0F0F0)  '设置绘制区域的背景色，也起清除作用
        Call TextOut(dc, 10, 100, pro, Len(CStr(pro))) '输出文本！
        Call Sleep(40) '等待
    Call ReleaseDC(formhandle, dc)  '释放资源！
   ' Call EnterCriticalSection(sect)
   ' 上下表示该处为临界区，如果要对工程全局变量做操作，最好在该区域内
   ' 否则线程同步过程中，非常容易让程序崩溃
   'Call LeaveCriticalSection(sect)
End Sub

Sub OutText2()  '和过程一类似
Dim i As Long
Dim dc As Long
Dim s As String
    dc = GetDC(formhandle)
    For i = 1 To 100000
        s = CStr(i)
        Call SetBkColor(dc, &HF0F0F0)
        Call TextOut(dc, 10, 80, s, Len(s))  '文本位置改变了
        Call Sleep(20) '延时改变了
    Next
    Call ReleaseDC(formhandle, dc)
   ' Call EnterCriticalSection(sect)
  '  Call LeaveCriticalSection(sect)
End Sub



'关于为何使用gdi 函数输出文本，这是一个很重要的内容；
'程序在记数时用了难用的TextOut 函数，而没有使用标签控件，这是因为
'vb的组件不都是线程安全的，当多线程访问不是线程安全的组件，那么会
'产生严重错误。
