VERSION 5.00
Begin VB.UserControl FCombo 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   ScaleHeight     =   390
   ScaleWidth      =   1785
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1695
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   1695
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   45
         TabIndex        =   3
         Top             =   30
         Width           =   135
      End
      Begin VB.Label LB1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   60
         Left            =   45
         TabIndex        =   0
         Top             =   30
         Width           =   570
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "FatCombo.ctx":0000
      Left            =   0
      List            =   "FatCombo.ctx":0002
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Combo1"
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   60
      Left            =   2280
      Picture         =   "FatCombo.ctx":0004
      Top             =   2520
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "FCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Const CB_SHOWDROPDOWN = &H14F

'事件声明:
Event Change() 'MappingInfo=Combo1,Combo1,-1,Change
Attribute Change.VB_Description = "当控件内容改变时发生。"
Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Event Expand()
Dim IfMove As Boolean
Dim EText As Boolean

Private Sub Combo1_Click()
On Error Resume Next
    Text1 = Combo1.Text
    PrintText
    If Text1.Visible = True Then
    Text1.SetFocus
    Else
    P1.SetFocus
    End If
    SetIt
    RaiseEvent Click
End Sub

Public Function SetDropdownWidth(NewWidthPixel As Long) As Boolean
    MoveWindow Combo1.hwnd, Combo1.Left, Combo1.Top, Combo1.Width / 15, NewWidthPixel, 1
End Function

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "返回一个 Font 对象。"
Attribute Font.VB_UserMemId = -512
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1.Font = New_Font
    Set LB1.Font = Combo1.Font
    Set Text1.Font = Combo1.Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "返回/设置对象中文本和图形的前景色。"
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Combo1.ForeColor() = New_ForeColor
    LB1.ForeColor = New_ForeColor
    Text1.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'背景颜色呢
Public Property Get BackColor() As OLE_COLOR
    BackColor = Combo1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Combo1.BackColor() = New_BackColor
    LB1.BackColor = New_BackColor
    Text1.BackColor = New_BackColor
    SetIt
    PropertyChanged "BackColor"
End Property

'''''''''''''
Public Property Get EnabledText() As Boolean
     EnabledText = Text1.Visible
End Property

Public Property Let EnabledText(ByVal New_EnabledText As Boolean)
    EText = New_EnabledText
    Text1.Visible = New_EnabledText
    PropertyChanged "EnabledText"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "返回/设置控件中包含的文本。"
    If Text1.Visible = True Then
        Text = Text1
    Else
        Text = LB1.caption
    End If
End Property

Public Property Let Text(ByVal New_Text As String)
    Text1 = New_Text
    Combo1.Text = New_Text
    PrintText
    PropertyChanged "Text"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "返回/设置控件的列表部分中包含的项。"
    List = Combo1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Combo1.List(Index) = New_List
    PropertyChanged "List"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "返回控件的列表部分中的项目数。"
    ListCount = Combo1.ListCount
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "返回/设置该控件中当前选定项目的索引。"
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
On Error GoTo aaaa
    Combo1.ListIndex() = New_ListIndex

    PropertyChanged "ListIndex"
aaaa:
End Property

Public Sub SetF()
On Error Resume Next
    Text1.SetFocus
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Private Sub PrintText()
On Error Resume Next
    LB1.caption = Text1.Text
    LB1.Top = (P1.Height - LB1.Height) / 2
    Text1.Top = (P1.Height - Text1.Height) / 2
    If LB1.Width > P1.Width - 300 Then LB1.Width = P1.Width - 300
End Sub


Private Sub Combo1_LostFocus()
    SetIt
End Sub

Private Sub LB1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
P1_MouseDown Button, Shift, x, y
End Sub



Private Sub LB1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
P1_MouseUp Button, Shift, x, y
End Sub


Private Sub P1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

IfMove = False
If Button = 1 Then
    xx = P1.Width - 15 - 225
    If x < xx Then Exit Sub
    P1.Line (xx - 15, 20)-(xx + 210, P1.Height - 30), 11899525, BF
    P1.PaintPicture Image1.Picture, xx - 15 + (230 - 105) / 2, (P1.Height - 60) / 2, 105, 60, 0, 0, 105, 60
End If

End Sub

Private Sub P1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

Dim MouseOver As Boolean
    '判断当前鼠标位置是否在控件上
MouseOver = (0 <= x) And (x <= P1.Width) And (0 <= y) And (y <= P1.Height)
If MouseOver Then
If IfMove = True Then Exit Sub
IfMove = True
SetIt2
SetCapture P1.hwnd
Else
IfMove = False
SetIt
ReleaseCapture
End If
End Sub

Private Sub SetIt2()
On Error Resume Next

P1.Line (0, 0)-(P1.Width - 15, P1.Height - 15), 6956042, B
'
xx = P1.Width - 30 - 225
P1.Line (xx - 15, 15)-(xx - 15, P1.Height - 30), 6956042, B
P1.Line (xx, 20)-(xx + 225, P1.Height - 30), 13811126, BF
P1.PaintPicture Image1.Picture, xx + (230 - 105) / 2, (P1.Height - 60) / 2, 105, 60, 0, 0, 105, 60

End Sub


Private Sub P1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    IfMove = False
    SetIt2
    If Button = 1 Then
        If Text1.Visible = True And x < P1.Width - 300 Then
            Text1.SetFocus
        Else
            RaiseEvent Expand
            Combo1.SetFocus
            SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
        End If
    End If
End Sub


Private Sub Text1_Change()
If Text1.Visible = True Then RaiseEvent Change
End Sub

Private Sub Text1_GotFocus()
Text1.SelLength = 1
Text1.SelLength = Len(Text1)
End Sub



Private Sub UserControl_Initialize()
    LB1.Top = (P1.Height - LB1.Height) / 2
    Text1.Top = (P1.Height - Text1.Height) / 2
End Sub

Public Sub SelAll()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'从存贮器中加载属性值
Public Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Dim Index As Integer

    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    LB1.ForeColor = Combo1.ForeColor
    Text1.ForeColor = Combo1.ForeColor
    Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Text1.BackColor = Combo1.BackColor
    SetIt
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    Text1.Visible = PropBag.ReadProperty("EnabledText", True)
    EText = PropBag.ReadProperty("EnabledText", True)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

UserControl.Height = Combo1.Height
P1.Height = Combo1.Height

If UserControl.Width < 390 Then UserControl.Width = 390
P1.Width = UserControl.Width
Combo1.Width = UserControl.Width
Text1.Width = UserControl.Width - 315
LB1.Top = (P1.Height - LB1.Height) / 2
Text1.Top = (P1.Height - Text1.Height) / 2
SetIt

End Sub

Public Sub SetCboWidth(ByVal cboWidth As Long)
    Combo1.Width = cboWidth
End Sub

Public Sub AddItem(Item, Optional Index)
On Error Resume Next

If IsMissing(Index) Then
Combo1.AddItem Item
Else
Combo1.AddItem Item, Index
End If
End Sub

Public Sub RemoveItem(Index)
Combo1.RemoveItem Index
End Sub

Public Sub SetIt()
On Error Resume Next

P1.Cls
P1.Line (0, 0)-(P1.Width - 15, P1.Height - 15), vbButtonFace, B
P1.Line (15, 15)-(P1.Width - 30, P1.Height - 30), Combo1.BackColor, BF
xx = P1.Width - 30 - 225
P1.Line (xx, 30)-(xx + 210, P1.Height - 45), vbButtonFace, BF
P1.PaintPicture Image1.Picture, xx + (230 - 105) / 2, (P1.Height - 60) / 2, 105, 60, 0, 0, 105, 60
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

Dim Index As Integer
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
    Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
    Call PropBag.WriteProperty("Text", Combo1.Text, "Combo1")
    Call PropBag.WriteProperty("EnabledText", EText, True)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)

End Sub

Public Sub Clear()
    Combo1.Clear
    Text1 = ""
    LB1.caption = ""
End Sub
