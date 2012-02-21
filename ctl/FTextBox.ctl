VERSION 5.00
Begin VB.UserControl FTextBox 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ScaleHeight     =   300
   ScaleWidth      =   2325
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1335
   End
   Begin VB.PictureBox Sh1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   15
      ScaleHeight     =   270
      ScaleWidth      =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   255
   End
End
Attribute VB_Name = "FTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Const CB_SHOWDROPDOWN = &H14F
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Dim IfOn As Boolean
Dim onX As Long
Dim onY As Long
Dim sX As Long
Dim sY As Long
'事件声明:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Text1,Text1,-1,KeyUp
Event Change() 'MappingInfo=Text1,Text1,-1,Change
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Text1,Text1,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Text1,Text1,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=Text1,Text1,-1,MouseUp
Event Click() 'MappingInfo=Text1,Text1,-1,Click
'缺省属性值:
Const m_def_afterdecimal = 2
Const m_def_isNumber = False
Const m_def_AutoSelAll = False
'属性变量:
Dim m_afterdecimal As Long
Dim m_isNumber As Boolean
Dim m_AutoSelAll As Boolean


Private Sub Text1_GotFocus()
On Error Resume Next
    If m_AutoSelAll Then
        Text1.SelStart = 0
        Text1.SelLength = Len(Text1)
    End If
End Sub

Private Sub Text1_LostFocus()
    UserControl.Cls
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 300
    Sh1.Width = UserControl.Width - 30
    Text1.Width = UserControl.Width - 90
End Sub
'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = Text1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Text1.BackColor() = New_BackColor
    Sh1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,Font
Public Property Get Font() As Font
    Set Font = Text1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Text1.Font = New_Font
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontBold
Public Property Get FontBold() As Boolean
    FontBold = Text1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Text1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = Text1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Text1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontName
Public Property Get FontName() As String
    FontName = Text1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Text1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontSize
Public Property Get FontSize() As Single
    FontSize = Text1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Text1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = Text1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Text1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = Text1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Text1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Text1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Text1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If m_isNumber Then
        If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then
        Else
            If m_afterdecimal = 0 And KeyAscii = 46 Then
                KeyAscii = 0
                Exit Sub
            End If
            If KeyAscii < 46 Or KeyAscii > 58 Or KeyAscii = 47 Then
                KeyAscii = 0
            Else
                If KeyAscii = 46 Then
                    If InStr(Text1.Text, ".") > 0 Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Text1.Text) - Text1.SelStart > m_afterdecimal Then
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
                im = InStr(Text1.Text, ".")
                If im > 0 Then
                    ix = Text1.SelStart
                    If im <= ix Then
                        If (ix - im) + 1 > m_afterdecimal Or Len(Text1.Text) + 1 - im > m_afterdecimal Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                Else
                    If KeyAscii = 46 And Len(Text1.Text) = (Text1.MaxLength - m_afterdecimal - 1) Then
                        Exit Sub
                    End If
                    If Text1.SelLength = Len(Text1.Text) And Text1.SelStart = 0 Then Exit Sub
                    If Len(Text1.Text) >= (Text1.MaxLength - m_afterdecimal - 1) Then KeyAscii = 0
                End If
                If Text1.SelLength = Len(Text1.Text) And Text1.SelStart = 0 Then Exit Sub
                If Len(Text1) >= Text1.MaxLength Then KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,Locked
Public Property Get Locked() As Boolean
    Locked = Text1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Text1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,Text
Public Property Get Text() As String
    Text = Text1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1.Text() = New_Text
    PropertyChanged "Text"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,SelLength
Public Property Get SelLength() As Long
    SelLength = Text1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Text1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,SelStart
Public Property Get SelStart() As Long
    SelStart = Text1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Text1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,SelText
Public Property Get SelText() As String
    SelText = Text1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Text1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

Private Sub Text1_Change()
    RaiseEvent Change
    If m_isNumber And Text1 = "" Then
        Text1 = "0"
        Text1.SelStart = 0
        Text1.SelLength = 1
    End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
    
    Dim pos As POINTAPI
    If IfOn = True Then Exit Sub
    onX = x / 15
    onY = y / 15
    GetCursorPos pos
    sX = pos.x
    sY = pos.y
    IfOn = True
    UserControl.Line (0, 0)-(UserControl.Width - 15, UserControl.Height - 15), 6956042, B
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Dim pos As POINTAPI, l As Long, t As Long, r As Long, b As Long
    GetCursorPos pos
    l = sX - onX
    t = sY - onY
    r = l + UserControl.Width / 15
    b = t + UserControl.Height / 15
    ScreenX = l
    ScreenY = t
    If pos.x < l Or pos.x > r Or pos.y < t Or pos.y > b Then
        IfOn = False
        UserControl.Cls
        Timer1.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Text1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set Text1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Text1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Text1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Text1.FontName = PropBag.ReadProperty("FontName", "")
    Text1.FontSize = PropBag.ReadProperty("FontSize", 0)
    Text1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Text1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Text1.Locked = PropBag.ReadProperty("Locked", False)
    Text1.Text = PropBag.ReadProperty("Text", "")
    Text1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Text1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Text1.SelText = PropBag.ReadProperty("SelText", "")
    Text1.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    m_AutoSelAll = PropBag.ReadProperty("AutoSelAll", m_def_AutoSelAll)
    Text1.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_isNumber = PropBag.ReadProperty("isNumber", m_def_isNumber)
    Text1.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    m_afterdecimal = PropBag.ReadProperty("afterdecimal", m_def_afterdecimal)
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", Text1.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", Text1.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", Text1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Text1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Text1.FontName, "")
    Call PropBag.WriteProperty("FontSize", Text1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", Text1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", Text1.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", Text1.Locked, False)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("SelLength", Text1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Text1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Text1.SelText, "")
    Call PropBag.WriteProperty("PasswordChar", Text1.PasswordChar, "")
    Call PropBag.WriteProperty("Enabled", Text1.Enabled, True)
    Call PropBag.WriteProperty("AutoSelAll", m_AutoSelAll, m_def_AutoSelAll)
    Call PropBag.WriteProperty("Alignment", Text1.Alignment, 0)
    Call PropBag.WriteProperty("isNumber", m_isNumber, m_def_isNumber)
    Call PropBag.WriteProperty("MaxLength", Text1.MaxLength, 0)
    Call PropBag.WriteProperty("afterdecimal", m_afterdecimal, m_def_afterdecimal)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,PasswordChar
Public Property Get PasswordChar() As String
    PasswordChar = Text1.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    Text1.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = Text1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get AutoSelAll() As Boolean
    AutoSelAll = m_AutoSelAll
End Property

Public Property Let AutoSelAll(ByVal New_AutoSelAll As Boolean)
    m_AutoSelAll = New_AutoSelAll
    PropertyChanged "AutoSelAll"
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_AutoSelAll = m_def_AutoSelAll
    m_isNumber = m_def_isNumber
    m_afterdecimal = m_def_afterdecimal
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = Text1.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    Text1.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,False
Public Property Get isNumber() As Boolean
    isNumber = m_isNumber
End Property

Public Property Let isNumber(ByVal New_isNumber As Boolean)
    m_isNumber = New_isNumber
    PropertyChanged "isNumber"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Text1,Text1,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = Text1.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    Text1.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=8,0,0,2
Public Property Get afterdecimal() As Long
    afterdecimal = m_afterdecimal
End Property

Public Property Let afterdecimal(ByVal New_afterdecimal As Long)
    m_afterdecimal = New_afterdecimal
    PropertyChanged "afterdecimal"
End Property


