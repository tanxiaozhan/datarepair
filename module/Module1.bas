Attribute VB_Name = "Module1"
'�뽫�ò������ݱ���Ϊ Module1.bas �ļ�

'�̰߳�ȫ�������ݽṹ��
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'��������ڶ��̷߳����ٽ���Դͬ��Api�����ݽṹ
Public Type CRITICAL_SECTION
    dummy As Long
End Type
'Ϊʲô��GDI ������ͼ��ԭ������ٽ�
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
'��ע�⣻createThread APi�����ѱ����޸Ĺ����޸ĵĵط������в���APIView���Ƶ�����
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
'�����sleep,���þ����������̻߳�ͼƵ�ʲ�һ�£�Ч�������ԡ�
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub EnterCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�����ٽ���
Public Declare Sub LeaveCriticalSection Lib "kernel32" (lpCriticalSection As CRITICAL_SECTION)  '�뿪�ٽ���

'������Ҫ�ĺ�������
'ObjPtr�����ض���ʵ��˽����ĵ�ַ��
'StrPtr�������ַ�����һ���ֵĵ�ַ��
'VarPtr�����ر����ĵ�ַ��

'ȫ�ֵ�form�ľ����
Public formhandle As Long
'�ٽ����ݽṹ
Public sect As CRITICAL_SECTION
Public pro As Long

Sub OutText1()  '����һ
Dim i As Long
Dim dc As Long
Dim s As String
    dc = GetDC(formhandle) '��ȡ��������DC
        Call SetBkColor(dc, &HF0F0F0)  '���û�������ı���ɫ��Ҳ���������
        Call TextOut(dc, 10, 100, pro, Len(CStr(pro))) '����ı���
        Call Sleep(40) '�ȴ�
    Call ReleaseDC(formhandle, dc)  '�ͷ���Դ��
   ' Call EnterCriticalSection(sect)
   ' ���±�ʾ�ô�Ϊ�ٽ��������Ҫ�Թ���ȫ�ֱ���������������ڸ�������
   ' �����߳�ͬ�������У��ǳ������ó������
   'Call LeaveCriticalSection(sect)
End Sub

Sub OutText2()  '�͹���һ����
Dim i As Long
Dim dc As Long
Dim s As String
    dc = GetDC(formhandle)
    For i = 1 To 100000
        s = CStr(i)
        Call SetBkColor(dc, &HF0F0F0)
        Call TextOut(dc, 10, 80, s, Len(s))  '�ı�λ�øı���
        Call Sleep(20) '��ʱ�ı���
    Next
    Call ReleaseDC(formhandle, dc)
   ' Call EnterCriticalSection(sect)
  '  Call LeaveCriticalSection(sect)
End Sub



'����Ϊ��ʹ��gdi ��������ı�������һ������Ҫ�����ݣ�
'�����ڼ���ʱ�������õ�TextOut ��������û��ʹ�ñ�ǩ�ؼ���������Ϊ
'vb������������̰߳�ȫ�ģ������̷߳��ʲ����̰߳�ȫ���������ô��
'�������ش���
