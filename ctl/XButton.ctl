VERSION 5.00
Begin VB.UserControl XButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1320
   DefaultCancel   =   -1  'True
   ScaleHeight     =   570
   ScaleWidth      =   1320
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   360
      Top             =   120
   End
   Begin VB.PictureBox im1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3120
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image img1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "XButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const m_def_IfDraw = False
Const m_def_style = 0
Const MouseOnC_def = vbMenuBar
Const MouseDownC_def = vbMenuBar
Const StyleC_def = 0
Const StyleC_def2 = -1
Const StyleC_3D1 = 16577259
Const StyleC_3D2 = 8421504
Dim IfOn As Boolean
Dim onX As Long
Dim onY As Long
Dim sX As Long
Dim sY As Long
Dim LeftClick As Long
Dim m_IfDraw As Boolean
Dim DownUpTime As Long
Public ScreenX As Long
Public ScreenY As Long
Enum Kstyle
    LeftPic
    TopPic
End Enum
Dim m_style As Kstyle
Dim MouseOnC As Long
Dim MouseDownC As Long
Dim StyleC As Long
Dim StyleC2 As Long
Dim NowC As Long
Dim Style3D1 As Long
Dim Style3D2 As Long
Event Click()
Event MouseOn()
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type color
    Red As Long
    Green As Long
    Blue As Long
End Type


Public Property Get IfDraw() As Boolean
    IfDraw = m_IfDraw
End Property

Public Property Let IfDraw(ByVal New_IfDraw As Boolean)
    m_IfDraw = New_IfDraw
    SetButton
    PropertyChanged "IfDraw"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    NowC = New_BackColor
    UserControl.BackColor = New_BackColor
    im1.BackColor = New_BackColor
    SetButton
    PropertyChanged "BackColor"
End Property

Public Property Get Tag() As String
    Tag = img1.Tag
End Property

Public Property Let Tag(ByVal New_Tag As String)
    img1.Tag = New_Tag
    PropertyChanged "Tag"
End Property

Public Property Get ToolTip() As String
    ToolTip = img1.ToolTipText
End Property

Public Property Let ToolTip(ByVal New_ToolTip As String)
    img1.ToolTipText = New_ToolTip
    PropertyChanged "ToolTip"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    SetButton
    PropertyChanged "ForeColor"
End Property

Public Property Get MouseOnColor() As OLE_COLOR
    MouseOnColor = MouseOnC
End Property

Public Property Let MouseOnColor(ByVal New_ForeColor As OLE_COLOR)
    MouseOnC = New_ForeColor
    SetButton
    PropertyChanged "MouseOnColor"
End Property

Public Property Get MouseDownColor() As OLE_COLOR
    MouseDownColor = MouseDownC
End Property

Public Property Let MouseDownColor(ByVal New_ForeColor As OLE_COLOR)
    MouseDownC = New_ForeColor
    SetButton
    PropertyChanged "MouseDownColor"
End Property

Public Property Get Style3DColor1() As OLE_COLOR
    Style3DColor1 = Style3D1
End Property

Public Property Let Style3DColor1(ByVal New_ForeColor As OLE_COLOR)
    Style3D1 = New_ForeColor
    SetButton
    PropertyChanged "Style3DColor1"
End Property

Public Property Get Style3DColor2() As OLE_COLOR
    Style3DColor2 = Style3D2
End Property

Public Property Let Style3DColor2(ByVal New_ForeColor As OLE_COLOR)
    Style3D2 = New_ForeColor
    SetButton
    PropertyChanged "Style3DColor2"
End Property

Public Property Get StyleColor() As OLE_COLOR
    StyleColor = StyleC
End Property

Public Property Let StyleColor(ByVal New_ForeColor As OLE_COLOR)
    StyleC = New_ForeColor
    SetButton
    PropertyChanged "StyleColor"
End Property

Public Property Get StyleColor2() As Long
    StyleColor2 = StyleC2
End Property

Public Property Let StyleColor2(ByVal New_ForeColor As Long)
    StyleC2 = New_ForeColor
    SetButton
    PropertyChanged "StyleColor2"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=im1,im1,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = im1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set im1.Picture = New_Picture
    SetButton
    PropertyChanged "Picture"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=14,0,0,0
Public Property Get style() As Kstyle
    style = m_style
End Property

Public Property Let style(ByVal New_style As Kstyle)
    m_style = New_style
    SetButton
    PropertyChanged "style"
End Property

Private Sub img1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub img1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub img1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseUp Button, Shift, x, y
End Sub

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_style = m_def_style
    MouseOnC = MouseOnC_def
    MouseDownC = MouseDownC_def
    StyleC = StyleC_def
    StyleC2 = StyleC_def2
    Style3D1 = StyleC_3D1
    Style3D2 = StyleC_3D2
    Set UserControl.Font = Ambient.Font
    m_IfDraw = m_def_IfDraw
    NowC = UserControl.BackColor
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Tag = PropBag.ReadProperty("Caption", "XButton")
    img1.ToolTipText = PropBag.ReadProperty("ToolTip", "")
    img1.Tag = PropBag.ReadProperty("Tag", "")
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    MouseDownC = PropBag.ReadProperty("MouseDownColor", &H80000012)
    MouseOnC = PropBag.ReadProperty("MouseOnColor", &H80000012)
    StyleC = PropBag.ReadProperty("StyleColor", &H80000012)
    StyleC2 = PropBag.ReadProperty("StyleColor2", -1)
    Style3D1 = PropBag.ReadProperty("Style3dColor1", &H80000012)
    Style3D2 = PropBag.ReadProperty("Style3dColor2", &H80000012)
    m_style = PropBag.ReadProperty("style", m_def_style)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_IfDraw = PropBag.ReadProperty("IfDraw", m_def_IfDraw)
    SetButton
End Sub

Private Sub UserControl_Show()
    NowC = UserControl.BackColor
    DrawMouseOut
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", UserControl.Tag, "XButton")
    Call PropBag.WriteProperty("ToolTip", img1.ToolTipText, "")
    Call PropBag.WriteProperty("Tag", img1.Tag, "")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseDownColor", MouseDownC, &H80000012)
    Call PropBag.WriteProperty("MouseOnColor", MouseOnC, &H80000012)
    Call PropBag.WriteProperty("StyleColor", StyleC, &H80000012)
    Call PropBag.WriteProperty("StyleColor2", StyleC2, -1)
    Call PropBag.WriteProperty("Style3dColor1", Style3D1, &H80000012)
    Call PropBag.WriteProperty("Style3dColor2", Style3D2, &H80000012)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("style", m_style, m_def_style)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("IfDraw", m_IfDraw, m_def_IfDraw)
End Sub

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    SetButton
End Property

Public Property Get caption() As String
    caption = UserControl.Tag
End Property

Public Property Let caption(ByVal New_caption As String)
    UserControl.Tag() = New_caption
    PropertyChanged "Caption"
End Property

'注意！不要删除或修改下列被注释的行！
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    SetButton
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DownUpTime = 1
    RaiseEvent MouseDown(Button, Shift, x, y)
    If Button = 2 Then
        IfOn = False
    End If
    LeftClick = Button
    If Button = 1 Then
        NowC = MouseDownC
        SetButton
        UserControl.Line (0, 0)-(0, UserControl.Height), Style3D2
        UserControl.Line (0, 0)-(UserControl.Width, 0), Style3D2
        UserControl.Line (UserControl.Width - 10, 0)-(UserControl.Width - 10, UserControl.Height), Style3D1
        UserControl.Line (0, UserControl.Height - 10)-(UserControl.Width, UserControl.Height - 10), Style3D1
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DownUpTime = 0
    RaiseEvent MouseUp(Button, Shift, x, y)
    DrawMouseOut
    If LeftClick = 1 And IfOn = True Then RaiseEvent Click
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pos As POINTAPI
    If IfOn = True Or DownUpTime = 1 Then Exit Sub
    RaiseEvent MouseOn
    onX = x / 15
    onY = y / 15
    GetCursorPos pos
    sX = pos.x
    sY = pos.y
    IfOn = True
    NowC = MouseOnC
    DrawMouseOn
    Timer1.Enabled = True
End Sub

Public Sub DrawMouseOn()
    SetButton
    UserControl.Line (0, 0)-(0, UserControl.Height), Style3D1
    UserControl.Line (0, 0)-(UserControl.Width, 0), Style3D1
    UserControl.Line (UserControl.Width - 10, 0)-(UserControl.Width - 10, UserControl.Height), Style3D2
    UserControl.Line (0, UserControl.Height - 10)-(UserControl.Width, UserControl.Height - 10), Style3D2
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
        RaiseEvent MouseOut
        IfOn = False
        DrawMouseOut
        Timer1.Enabled = False
        Exit Sub
    End If
    If DownUpTime = 0 Then DrawMouseOn
End Sub

Private Sub UserControl_Resize()
    img1.Move 0, 0, UserControl.Width, UserControl.Height
    SetButton
End Sub

Public Sub DrawMouseOut()
    NowC = UserControl.BackColor
    UserControl.Line (0, 0)-(UserControl.Width - 10, UserControl.Height - 10), UserControl.BackColor, B
    SetButton
End Sub

'下面打印按钮
Public Sub SetButton()
On Error Resume Next
    x = im1.Width: y = im1.Height
    If NowC = 0 Then NowC = UserControl.BackColor
    UserControl.Line (0, 0)-(UserControl.Width, UserControl.Height), NowC, BF
    
    If m_IfDraw = True Then
        If StyleC2 = -1 Then
            UserControl.Line (0, 0)-(UserControl.Width - 10, UserControl.Height - 10), StyleC, B
        Else
            UserControl.Line (0, 0)-(0, UserControl.Height), StyleC
            UserControl.Line (0, 0)-(UserControl.Width, 0), StyleC
            UserControl.Line (UserControl.Width - 10, 0)-(UserControl.Width - 10, UserControl.Height), StyleC2
            UserControl.Line (0, UserControl.Height - 10)-(UserControl.Width, UserControl.Height - 10), StyleC2
        End If
    End If
    If caption = "" And im1.Picture = LoadPicture() Then Exit Sub
    If caption = "" Then PrintMePicture (UserControl.Width - x) / 2, (UserControl.Height - y) / 2: Exit Sub
    If im1.Picture = LoadPicture("") Then
        PrintMeCaption (UserControl.Width - TextWidth(caption)) / 2, (UserControl.Height - TextHeight(caption)) / 2
        Exit Sub
    End If
    If m_style = 0 Then
        PrintMePicture (UserControl.Width - x - TextWidth(caption)) / 3, (UserControl.Height - y) / 2
        PrintMeCaption x + 2 * (UserControl.Width - x - TextWidth(caption)) / 3, (UserControl.Height - TextHeight(caption)) / 2
        Exit Sub
    End If
    If m_style = 1 Then
        PrintMePicture (UserControl.Width - x) / 2, (UserControl.Height - y - TextHeight(caption)) / 3
        PrintMeCaption (UserControl.Width - TextWidth(caption)) / 2, y + 2 * (UserControl.Height - TextHeight(caption) - y) / 3
    End If
End Sub

Private Function GrayScaleColor(color) As Long
    Dim ColorValues As color
    ColorValues = RGBValues(color)
    With ColorValues
        GrayScaleColor = (9798 * .Red + 19235 * .Green + 3735 * .Blue) \ 32768
        .Red = GrayScaleColor
        .Green = GrayScaleColor
        .Blue = GrayScaleColor
        GrayScaleColor = RGB(.Red, .Green, .Blue)
    End With
End Function

Private Function RGBValues(color) As color  'find the rgb color values of a color
    Dim ReturnColor As color
    With ReturnColor
        .Red = Fix(color And 255)
        .Green = Fix((color And 65535) / 256)
        .Blue = Fix(color / 65536)
    End With
    RGBValues = ReturnColor
End Function

Private Sub PrintMePicture(ByVal x As Long, ByVal y As Long)
    If UserControl.Enabled Then
        UserControl.PaintPicture im1.Picture, x, y, im1.Width, im1.Height
    Else
        Dim i As Long, j As Long, n As Long, n2 As Long
        n = UserControl.Point(0, 0)
        For i = 0 To im1.Width - 15 Step 15
            For j = 0 To im1.Height - 15 Step 15
                n2 = im1.Point(i, j)
                If n2 <> n Then UserControl.PSet (x + i, y + j), GrayScaleColor(n2)
            Next
        Next
    End If
End Sub

Private Sub PrintMeCaption(ByVal x As Long, ByVal y As Long)
    If UserControl.Enabled Then
        UserControl.CurrentX = x
        UserControl.CurrentY = y
        UserControl.Print UserControl.Tag
    Else
        Dim j As Long
        j = UserControl.ForeColor
        UserControl.ForeColor = 16777215
        UserControl.CurrentX = x + 15
        UserControl.CurrentY = y + 15
        UserControl.Print UserControl.Tag
        UserControl.ForeColor = 8421504
        UserControl.CurrentX = x
        UserControl.CurrentY = y
        UserControl.Print UserControl.Tag
        UserControl.ForeColor = j
    End If
End Sub
