VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "�����޸�"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   StartUpPosition =   2  '��Ļ����
   Begin DataRepair.XPButton XPButton2 
      Height          =   615
      Left            =   2445
      TabIndex        =   2
      Top             =   3480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "�˳��޸�����"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
      Height          =   615
      Left            =   2445
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      Caption         =   "ִ�������޸�"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
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
   Begin MSComctlLib.ProgressBar ProBar 
      Height          =   375
      Left            =   945
      TabIndex        =   3
      Top             =   2205
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   75
      TabIndex        =   7
      Top             =   3930
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label3 
      Caption         =   "�ܽ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   5
      Top             =   2265
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   1455
      Visible         =   0   'False
      Width           =   6675
   End
   Begin VB.Label Label1 
      Caption         =   "��ȷ�����û�ʹ�òֹ�ϵͳʱ��ִ�������޸����ܣ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1695
      TabIndex        =   1
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1275
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public btime As Variant
Dim h1 As Integer
Dim m1 As Integer



Private Sub XPButton1_Click()
    Dim n As Long
    Dim pro1 As Long
    Dim rs As adodb.Recordset
    Dim instorers As adodb.Recordset
    Dim outstoreRS As adodb.Recordset
    Dim bzh(5) As String
    h1 = 60 * 60
    m1 = 60
    
    bzh(0) = "һ"
    bzh(1) = "��"
    bzh(2) = "��"
    bzh(3) = "��"
    bzh(4) = "��"
    If MsgBox("ȷ�Ͻ��������޸���", vbExclamation + vbYesNo, "��ʾ") = vbNo Then Exit Sub
    btime = Timer()
    
    dispTime
    
    Label1.Visible = False
    Label3.Visible = True
    Label4.Visible = True
    XPButton1.Enabled = False
    XPButton2.Enabled = False
    Label1.Refresh
    Label4.caption = ""
    Label5.caption = ""
    Label5.Visible = True
    
    
    ProBar.Min = 0
    DBConnect
    
    Set rs = New adodb.Recordset
    rs.Open "select count(*) from instore", Conn, 1, 1
    pro1 = CLng(rs(0) * 0.1)
    ProBar.Max = pro1 + rs(0)
    rs.Close
    rs.Open "select count(*) from outstore", Conn, 1, 1
    ProBar.Max = ProBar.Max + rs(0) + rs(0)
    rs.Close
    
    ProBar.Visible = True
    Label2.Visible = True
    bz = 0
    Label2.caption = "�޸�������֮" & bzh(bz) & "  ��ʼ�����ֽ�������!"
    Label2.Refresh
    
    sql = "Update InStore " & _
          "set remainnumber = allnumber," & _
          "remainpack = packnumber," & _
          "remainweight=weight"
    Conn.Execute sql
    ProBar.value = ProBar.value + pro1
    ProBar.Refresh
    Label4.caption = Int(ProBar.value / ProBar.Max * 100) & "%"
    Label4.Refresh
    Label2.caption = "��ʼ�����ֽ����������!"
    Label2.Refresh
    
    dispTime
    
    bz = bz + 1
    
    Dim lRecno(1000) As Long
    n = 0
    inrecount = 0
    Set instorers = New adodb.Recordset
    Set outstoreRS = New adodb.Recordset
    sql = "select recno,InstoreRecNo from outstore"
    outstoreRS.Open sql, Conn, 1, 1
    total = outstoreRS.RecordCount
    Do While Not outstoreRS.EOF

        sql = "select count(*) from instore where recno = " & CStr(outstoreRS("instoreRecNo"))
        instorers.Open sql, Conn, 1, 1
        If instorers(0) < 1 Then
            n = n + 1
            lRecno(n) = outstoreRS("instoreRecNo")
        End If
        instorers.Close
        outstoreRS.MoveNext
        ProBar.value = ProBar.value + 1
        ProBar.Refresh
        Label4.caption = Int(ProBar.value / ProBar.Max * 100) & "%"
        Label4.Refresh
        
        Label2.caption = "�޸�������֮" & bzh(bz) & "  �����޸�������������!  " & ProBar.value & "/" & total
        Label2.Refresh
        
        dispTime
    
    Loop
    
    outstoreRS.Close
    bz = bz + 1
    For i = 1 To n
        sql = "delete from outstore where instoreRecNo=" & CStr(lRecno(i))
        Conn.Execute sql
        Label2.caption = "�޸�������֮" & bzh(bz) & "  ���������������!    " & i & "/" & n
        Label2.Refresh
        
        dispTime
        
        
    Next
    Label2.caption = "�޸������������ݳɹ�!"
    Label2.Refresh
    '�ý������ݿ��е����ݸ��³��ֿ������
    bz = bz + 1
    Label2.caption = "׼�����³�������!"
    Label2.Refresh
    sql = "select * from instore"
    instorers.Open sql, Conn, 1, 1
    total = instorers.RecordCount
    n = 0

    Do While Not instorers.EOF
        n = n + 1
        sql = "Update OutStore " & _
              "set ClientID = '" & instorers("ClientID") & "'," & _
              "contractID='" & instorers("contractID") & "'," & _
              "Gross=" & CStr(instorers("Gross")) & "," & _
              "net=" & CStr(instorers("net")) & "," & _
              "PackWeight=" & CStr(instorers("PackWeight")) & " " & _
              "where instoreRecno=" & instorers("RecNo")
              

        Conn.Execute sql
        
        '����δ������,����=����*�ܼ���
        sql = "Update OutStore " & _
              "set weight=pack1*" & CStr(instorers("packweight")) & "/1000  " & _
              "where instoreRecno=" & instorers("RecNo") & " and lockweight=FALSE"

        Conn.Execute sql
        
        instorers.MoveNext
        Label2.caption = "�޸�������֮" & bzh(bz) & "  ���ڸ��³������ݣ�  " & n & "/" & total
        ProBar.value = ProBar.value + 1
        ProBar.Refresh
        Label4.caption = Int(ProBar.value / ProBar.Max * 100) & "%"
        Label4.Refresh
        
        dispTime
        
    Loop
    
    Label2.caption = "���³������ݳɹ�!"
    Label2.Refresh


    '���ݳ�������,�޸����ֽ�������
    sql = "select * from OutStore"
    rs.Open sql, Conn, 1, 1
    bz = bz + 1
    Label2.caption = "׼���޸����ֽ�������!"
    Label2.Refresh
    total = rs.RecordCount
    n = 0
    
    Do While Not rs.EOF
       n = n + 1
        sql = "Update InStore " & _
              "set remainnumber = remainnumber - " & rs("number1") & "," & _
              "remainpack = remainpack - " & rs("pack1") & "," & _
              "remainweight=remainweight - " & rs("weight") & " " & _
              "where Recno=" & rs("InStoreRecNo")
        Conn.Execute sql
    
        rs.MoveNext
        ProBar.value = ProBar.value + 1
        ProBar.Refresh
        Label4.caption = Int(ProBar.value / ProBar.Max * 100) & "%"
        Label4.Refresh
        Label2.caption = "�޸�������֮" & bzh(bz) & "  �����޸���������!   " & n & "/" & total
        Label2.Refresh
        
       dispTime
       
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Label2.caption = "�޸����ֽ������ݳɹ�!"
    Label2.Refresh
    ProBar.value = ProBar.Max
    ProBar.Refresh
    Label4.caption = Int(ProBar.value / ProBar.Max * 100) & "%"
    Label4.Refresh

    Label2.caption = "�����޸����!"
    Label2.Refresh
    dispTime
    XPButton2.Enabled = True

    
    
End Sub

Private Sub XPButton2_Click()
    Unload Me
End Sub

Private Sub dispTime()
    s = Timer() - btime
    h = Int(s / h1)
    m = Int((s Mod h1) / m1)
    s = Int((s Mod h1) Mod m1)
    Label5.caption = "����ʱ�䣺" & h & ":" & Format(m, "0#") & ":" & Format(s, "0#")
    Label5.Refresh
    DoEvents
'    If GetForegroundWindow() <> Me.hwnd Then
'       MsgBox "����", vbInformation, "�����޸�"
'  End If
    
End Sub
